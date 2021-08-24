Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmLabelDesign
	Inherits System.Windows.Forms.Form
	'===============================================================================
	'  프로그램 : 오스템 임풀란트 메인 폼 [바코드레이아웃 불러오기/수정/저장,동적 컨트롤 생성/이벤트 처리]
	'  파 일 명 : frmLabelDesign.frm
	'  작 성 일 : 2011.09.21
	'  작 성 자 : 오세원
	'  홈페이지 : http://www.didiminfoinfo.co.kr
	'  설    명 :
	'  수정이력 :
	'===============================================================================
	
	
	Private m_ColCommandButton As Collection ' 동적 생성 컨트롤 저장을 위한 컬렉션
	Private WithEvents ClsEventMonitor As ClassEventMonitor ' 이벤트 전달을 위한 클래스
	
	'UPGRADE_WARNING: LOGFONT 구조체에는 이 Declare 문에 인수로 전달할 마샬링 특성이 있어야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function CreateFontIndirect Lib "gdi32"  Alias "CreateFontIndirectA"(ByRef lpLogFont As LOGFONT) As Integer
	Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Integer, ByVal hObject As Integer) As Integer
	Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Integer, ByVal nIndex As Integer) As Integer
	Private Declare Function TextOut Lib "gdi32"  Alias "TextOutA"(ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer, ByVal lpString As String, ByVal nCount As Integer) As Integer
	
	
	'==== API 파일 오픈 관련 =================================================
	Const FW_NORMAL As Short = 400
	Const DEFAULT_CHARSET As Short = 1
	Const OUT_DEFAULT_PRECIS As Short = 0
	Const CLIP_DEFAULT_PRECIS As Short = 0
	Const DEFAULT_QUALITY As Short = 0
	Const DEFAULT_PITCH As Short = 0
	Const FF_ROMAN As Short = 16
	Const CF_PRINTERFONTS As Integer = &H2
	Const CF_SCREENFONTS As Integer = &H1
	Const CF_BOTH As Boolean = (CF_SCREENFONTS Or CF_PRINTERFONTS)
	Const CF_EFFECTS As Integer = &H100
	Const CF_FORCEFONTEXIST As Integer = &H10000
	Const CF_INITTOLOGFONTSTRUCT As Integer = &H40
	Const CF_LIMITSIZE As Integer = &H2000
	Const REGULAR_FONTTYPE As Integer = &H400
	Const LF_FACESIZE As Short = 32
	Const CCHDEVICENAME As Short = 32
	Const CCHFORMNAME As Short = 32
	Const GMEM_MOVEABLE As Integer = &H2
	Const GMEM_ZEROINIT As Integer = &H40
	Const DM_DUPLEX As Integer = &H1000
	Const DM_ORIENTATION As Integer = &H1
	Const PD_PRINTSETUP As Integer = &H40
	Const PD_DISABLEPRINTTOFILE As Integer = &H80000
	
	Private Const ANSI_CHARSET As Short = 0
	Private Const VARIABLE_PITCH As Short = 2
	Private Const FF_DONTCARE As Short = 0
	Private Const FW_BOLD As Short = 700
	Private Const LOGPIXELSY As Short = 90
	
	
	'Private Type POINTAPI
	'    x As Long
	'    y As Long
	'End Type
	Private Structure RECT
		'UPGRADE_NOTE: Left이(가) Left_Renamed(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Left_Renamed As Integer
		'UPGRADE_NOTE: Top이(가) Top_Renamed(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Top_Renamed As Integer
		'UPGRADE_NOTE: Right이(가) Right_Renamed(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Right_Renamed As Integer
		'UPGRADE_NOTE: Bottom이(가) Bottom_Renamed(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Bottom_Renamed As Integer
	End Structure
	Private Structure OPENFILENAME
		Dim lStructSize As Integer
		Dim hwndOwner As Integer
		Dim hInstance As Integer
		Dim lpstrFilter As String
		Dim lpstrCustomFilter As String
		Dim nMaxCustFilter As Integer
		Dim nFilterIndex As Integer
		Dim lpstrFile As String
		Dim nMaxFile As Integer
		Dim lpstrFileTitle As String
		Dim nMaxFileTitle As Integer
		Dim lpstrInitialDir As String
		Dim lpstrTitle As String
		Dim flags As Integer
		Dim nFileOffset As Short
		Dim nFileExtension As Short
		Dim lpstrDefExt As String
		Dim lCustData As Integer
		Dim lpfnHook As Integer
		Dim lpTemplateName As String
	End Structure
	Private Structure PAGESETUPDLG
		Dim lStructSize As Integer
		Dim hwndOwner As Integer
		Dim hDevMode As Integer
		Dim hDevNames As Integer
		Dim flags As Integer
		Dim ptPaperSize As POINTAPI
		Dim rtMinMargin As RECT
		Dim rtMargin As RECT
		Dim hInstance As Integer
		Dim lCustData As Integer
		Dim lpfnPageSetupHook As Integer
		Dim lpfnPagePaintHook As Integer
		Dim lpPageSetupTemplateName As String
		Dim hPageSetupTemplate As Integer
	End Structure
	Private Structure CHOOSECOLOR
		Dim lStructSize As Integer
		Dim hwndOwner As Integer
		Dim hInstance As Integer
		Dim rgbResult As Integer
		Dim lpCustColors As String
		Dim flags As Integer
		Dim lCustData As Integer
		Dim lpfnHook As Integer
		Dim lpTemplateName As String
	End Structure
	Private Structure LOGFONT
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
		<VBFixedString(31),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=31)> Public lfFaceName() As Char
	End Structure
	Private Structure CHOOSEFONT
		Dim lStructSize As Integer
		Dim hwndOwner As Integer '  caller's window handle
		Dim hdc As Integer '  printer DC/IC or NULL
		Dim lpLogFont As Integer '  ptr. to a LOGFONT struct
		Dim iPointSize As Integer '  10 * size in points of selected font
		Dim flags As Integer '  enum. type flags
		Dim rgbColors As Integer '  returned text color
		Dim lCustData As Integer '  data passed to hook fn.
		Dim lpfnHook As Integer '  ptr. to hook function
		Dim lpTemplateName As String '  custom template name
		Dim hInstance As Integer '  instance handle of.EXE that
		'    contains cust. dlg. template
		Dim lpszStyle As String '  return the style field here
		'  must be LF_FACESIZE or bigger
		Dim nFontType As Short '  same value reported to the EnumFonts
		'    call back with the extra FONTTYPE_
		'    bits added
		Dim MISSING_ALIGNMENT As Short
		Dim nSizeMin As Integer '  minimum pt size allowed &
		Dim nSizeMax As Integer '  max pt size allowed if
		'    CF_LIMITSIZE is used
	End Structure
	Private Structure PRINTDLG_TYPE
		Dim lStructSize As Integer
		Dim hwndOwner As Integer
		Dim hDevMode As Integer
		Dim hDevNames As Integer
		Dim hdc As Integer
		Dim flags As Integer
		Dim nFromPage As Short
		Dim nToPage As Short
		Dim nMinPage As Short
		Dim nMaxPage As Short
		Dim nCopies As Short
		Dim hInstance As Integer
		Dim lCustData As Integer
		Dim lpfnPrintHook As Integer
		Dim lpfnSetupHook As Integer
		Dim lpPrintTemplateName As String
		Dim lpSetupTemplateName As String
		Dim hPrintTemplate As Integer
		Dim hSetupTemplate As Integer
	End Structure
	Private Structure DEVNAMES_TYPE
		Dim wDriverOffset As Short
		Dim wDeviceOffset As Short
		Dim wOutputOffset As Short
		Dim wDefault As Short
		'UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(100),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=100)> Public extra() As Char
	End Structure
	Private Structure DEVMODE_TYPE
		'UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(CCHDEVICENAME),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=CCHDEVICENAME)> Public dmDeviceName() As Char
		Dim dmSpecVersion As Short
		Dim dmDriverVersion As Short
		Dim dmSize As Short
		Dim dmDriverExtra As Short
		Dim dmFields As Integer
		Dim dmOrientation As Short
		Dim dmPaperSize As Short
		Dim dmPaperLength As Short
		Dim dmPaperWidth As Short
		Dim dmScale As Short
		Dim dmCopies As Short
		Dim dmDefaultSource As Short
		Dim dmPrintQuality As Short
		Dim dmColor As Short
		Dim dmDuplex As Short
		Dim dmYResolution As Short
		Dim dmTTOption As Short
		Dim dmCollate As Short
		'UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(CCHFORMNAME),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=CCHFORMNAME)> Public dmFormName() As Char
		Dim dmUnusedPadding As Short
		Dim dmBitsPerPel As Short
		Dim dmPelsWidth As Integer
		Dim dmPelsHeight As Integer
		Dim dmDisplayFlags As Integer
		Dim dmDisplayFrequency As Integer
	End Structure
	'UPGRADE_WARNING: CHOOSECOLOR 구조체에는 이 Declare 문에 인수로 전달할 마샬링 특성이 있어야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function CHOOSECOLOR_Renamed Lib "comdlg32.dll"  Alias "ChooseColorA"(ByRef pChoosecolor As CHOOSECOLOR) As Integer
	'UPGRADE_WARNING: OPENFILENAME 구조체에는 이 Declare 문에 인수로 전달할 마샬링 특성이 있어야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetOpenFileName Lib "comdlg32.dll"  Alias "GetOpenFileNameA"(ByRef pOpenfilename As OPENFILENAME) As Integer
	'UPGRADE_WARNING: OPENFILENAME 구조체에는 이 Declare 문에 인수로 전달할 마샬링 특성이 있어야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetSaveFileName Lib "comdlg32.dll"  Alias "GetSaveFileNameA"(ByRef pOpenfilename As OPENFILENAME) As Integer
	'UPGRADE_WARNING: PRINTDLG_TYPE 구조체에는 이 Declare 문에 인수로 전달할 마샬링 특성이 있어야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function PrintDialog Lib "comdlg32.dll"  Alias "PrintDlgA"(ByRef pPrintdlg As PRINTDLG_TYPE) As Integer
	'UPGRADE_WARNING: PAGESETUPDLG 구조체에는 이 Declare 문에 인수로 전달할 마샬링 특성이 있어야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function PAGESETUPDLG_Renamed Lib "comdlg32.dll"  Alias "PageSetupDlgA"(ByRef pPagesetupdlg As PAGESETUPDLG) As Integer
	'UPGRADE_WARNING: CHOOSEFONT 구조체에는 이 Declare 문에 인수로 전달할 마샬링 특성이 있어야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function CHOOSEFONT_Renamed Lib "comdlg32.dll"  Alias "ChooseFontA"(ByRef pChoosefont As CHOOSEFONT) As Integer
	'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
	Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Integer) As Integer
	Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Integer) As Integer
	Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Integer, ByVal dwBytes As Integer) As Integer
	Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Integer) As Integer
	
	Dim OFName As OPENFILENAME
	Dim CustomColors() As Byte
	'==== API 파일 오픈 관련 =================================================
	
	'Public Function DrawRotatedText(lhDC As Long, FontInfo As StdFont, iRot As Integer, sText As String, lX As Long, lY As Long) As Boolean
	'
	''On Error GoTo DrawRotatedText_E
	'
	''Parameters:
	''   lhDC     - The device context to draw the text on
	''   FontInfo - A font structure with the font to use
	''   iRot     - Rotation in tenths of degrees (900 equals 90 degrees)
	''   sText    - The text to draw
	''   lX       - X coordinate of starting point (in pixels)
	''   lY       - Y coordinate of starting point (in pixels)
	''
	''Return value:
	''   returns true if successful, false otherwise
	''
	''Last modified: June 9, 1999
	''Special thanks to: Sebastian Strand
	'
	'Dim hlFont As Long, hlOld As Long
	'Dim uLogFont As LOGFONT, b As Byte
	'Dim abChars() As Byte
	'
	''Fill logfont structure with proper font data
	'With uLogFont
	'
	'.lfCharSet = ANSI_CHARSET
	'.lfClipPrecision = CLIP_DEFAULT_PRECIS
	'.lfEscapement = iRot
	'
	''We can't assign directly to fixed length array
	''so we have to use a temp array and copy the chars
	''one by one
	'abChars = StrConv(FontInfo.Name, vbFromUnicode)
	'For b = 0 To IIf(UBound(abChars) > UBound(.lfFaceName), UBound(.lfFaceName), UBound(abChars))
	'.lfFaceName(b) = abChars(b)
	'Next b
	'
	'.lfHeight = FontInfo.Size / 72 * GetDeviceCaps(lhDC, LOGPIXELSY)
	'.lfWidth = 0 'When zero windows calculates proper width based on the height setting
	'.lfItalic = Abs(FontInfo.Italic)
	'.lfOrientation = .lfEscapement
	'.lfOutPrecision = OUT_DEFAULT_PRECIS
	'.lfPitchAndFamily = VARIABLE_PITCH Or FF_DONTCARE
	'.lfQuality = DEFAULT_QUALITY
	'.lfStrikeOut = Abs(FontInfo.Strikethrough)
	'.lfUnderline = Abs(FontInfo.Underline)
	'.lfWeight = IIf(FontInfo.Bold, FW_BOLD, FW_NORMAL)
	'End With
	'
	''Create font
	'hlFont = CreateFontIndirect(uLogFont)
	'If hlFont = 0 Then Exit Function
	'
	''Select created font into dc to use it
	'hlOld = SelectObject(lhDC, hlFont)
	'
	''Draw text and return result
	'DrawRotatedText = (TextOut(lhDC, lX, lY, sText, Len(sText)) <> 0)
	'
	''Select old font back
	'hlOld = SelectObject(lhDC, hlOld)
	'
	'DrawRotatedText_X:
	'Exit Function
	'
	'DrawRotatedText_E:
	'Resume DrawRotatedText_X
	'
	'End Function
	
	Private Sub ActiveResize1_BeforeResize(ByRef Cancel As Boolean)
		'
		'    Dim varBuffer() As Variant
		'    Dim varBuf      As Variant
		'    Dim utf8() As Byte
		'    Dim ucs2 As Variant
		'    Dim chars As Long
		'    Dim varTmp As Variant
		'    Dim i As Integer
		'    Dim LineCount As Long
		'
		'    If gOpenFileNm <> "" Then
		'        '컬렉션 초기화
		'        Set m_ColCommandButton = Nothing
		'        Set m_ColCommandButton = New Collection
		'
		'        gblCtrlNm = "Control_0"
		'        gblCtrlIdx = 0
		'
		'        Open gOpenFileNm For Binary As #1   'UTF-8 문서지정
		'        ReDim utf8(LOF(1))
		'
		'        Get #1, , utf8
		'
		'        chars = MultiByteToWideChar(CP_UTF8, 0, VarPtr(utf8(0)), LOF(1), 0, 0)
		'        ucs2 = Space(chars)
		'        chars = MultiByteToWideChar(CP_UTF8, 0, VarPtr(utf8(0)), LOF(1), StrPtr(ucs2), chars)
		'        varBuf = Split(ucs2, Chr(13))
		'
		'
		'        Close #1
		'
		'
		'        '오픈한 LOF파일 버퍼에 쓰기
		'        For i = 0 To UBound(varBuf)
		'            ReDim Preserve varBuffer(i)
		'            varBuffer(LineCount) = varBuf(i)
		'            LineCount = LineCount + 1
		'        Next
		'
		'
		'        '오픈한 LOF파일 화면그리기/스프레드쓰기
		'        For i = 0 To UBound(varBuffer) - 1
		'            If varBuffer(i) <> "" Then
		'                varBuf = Split(varBuffer(i), "^")
		'                Call MakeLayout(varBuf)
		'                Call SetList(varBuf)
		'            End If
		'        Next
		'
		'        Call PaintLine
		'    End If
		
	End Sub
	
	'UPGRADE_WARNING: 폼이 초기화될 때 cboType.SelectedIndexChanged 이벤트가 발생합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboType.SelectedIndexChanged
		
		sstType.SelectedIndex = cboType.SelectedIndex
		
		Select Case cboType.SelectedIndex
			Case 0
				txtTitle.Text = "S_TEXT" & gblCtrlIdx
			Case 1
				txtTitle.Text = "D_TEXT" & gblCtrlIdx
			Case 2
				txtTitle.Text = "S_Image" & gblCtrlIdx
			Case 3
				txtTitle.Text = "D_Image" & gblCtrlIdx
			Case 4
				txtTitle.Text = "BARCODE" & gblCtrlIdx
			Case 5
				txtTitle.Text = "LINE" & gblCtrlIdx
				txtLineHSize.Text = "1"
		End Select
		
		txtXpos.Text = CStr(1)
		txtYpos.Text = CStr(10)
		
	End Sub
	
	
	
	
	
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' 동적 생성 컨트롤에서의 이벤트 처리
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	Private Sub ClsEventMonitor_EventRaised(ByRef EventObject As ClassEventObject, ByVal StrEventName As String) Handles ClsEventMonitor.EventRaised
		
		Dim StrEvent As String
		'FIXIT: 'obj'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                               FixIT90210ae-R1672-R1B8ZE
		Dim obj As Object
		'FIXIT: 'val1'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                              FixIT90210ae-R1672-R1B8ZE
		Dim val1 As Object
		
		On Error Resume Next
		
		' 실제 이벤트가 발생한 Object
		obj = EventObject.EventObject
		
		StrEvent = ""
		StrEvent = StrEvent & VB6.Format(Now, "HH:MM:SS") & " "
		'UPGRADE_WARNING: obj.Name 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		StrEvent = StrEvent & obj.Name & " - " & StrEventName & "("
		
		' 파라미터 정보
		For	Each val1 In EventObject.Params
			'UPGRADE_WARNING: val1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			StrEvent = StrEvent & CStr(val1) & ", "
		Next val1
		
		'FIXIT: 'Right' 함수를 'Right$' 함수로 바꾸십시오.                                                    FixIT90210ae-R9757-R1B8ZE
		If VB.Right(StrEvent, 2) = ", " Then
			'FIXIT: 'Left' 함수를 'Left$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
			StrEvent = VB.Left(StrEvent, Len(StrEvent) - 2)
		End If
		
		StrEvent = StrEvent & "" & ")"
		
		' 이벤트 로그
		List1.Items.Insert(0, StrEvent)
		
	End Sub
	
	Private Sub cmdDelobj_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDelobj.Click
		Dim intRow As Short
		'FIXIT: 'strObjType'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                        FixIT90210ae-R1672-R1B8ZE
		Dim strObjType As Object
		'FIXIT: 'strObjName'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                        FixIT90210ae-R1672-R1B8ZE
		Dim strObjName As Object
		'FIXIT: 'strObjRotate'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                      FixIT90210ae-R1672-R1B8ZE
		Dim strObjRotate As Object
		
		CType(Me.Controls(txtTag.Text), Object).Visible = False
		
		Dim counter As Short
		With spdList
			counter = .MaxRows
			For intRow = 1 To counter
				.Row = intRow
				Call .GetText(2, intRow, strObjType)
				Call .GetText(28, intRow, strObjName)
				'
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				'UPGRADE_WARNING: strObjName 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: strObjType 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If strObjType = sstType.SelectedIndex And strObjName = Trim(txtTag.Text) Then
					.Action = FPSpread.ActionConstants.ActionDeleteRow
					.MaxRows = .MaxRows - 1
					Exit For
				End If
			Next 
		End With
		
	End Sub
	
	Private Sub cmdDevide_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDevide.Click
		Dim Index As Short = cmdDevide.GetIndex(eventSender)
		Dim intRow As Short
		Dim intCol As Short
		Dim strBuf() As String
		
		intMode = 2
		
		If Index = 0 Then
			If txtDevide.Text = "0.2" Then
				txtDevide.Text = "0.2"
			Else
				txtDevide.Text = CStr(CDbl(txtDevide.Text) - 0.2)
			End If
		Else
			txtDevide.Text = CStr(CDbl(txtDevide.Text) + 0.2)
		End If
		gDevide = txtDevide.Text
		
		' 컬렉션 초기화
		'UPGRADE_NOTE: m_ColCommandButton 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_ColCommandButton = Nothing
		m_ColCommandButton = New Collection
		
		With spdList
			sstType.Visible = False
			For intRow = 1 To .MaxRows
				.Row = intRow
				.Col = 1
				Erase strBuf
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				If Trim(.Text) <> "" Then
					ReDim Preserve strBuf(.MaxCols)
					For intCol = 1 To .MaxCols
						.Col = intCol
						'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
						strBuf(intCol - 1) = Trim(.Text)
					Next 
					Call MakeLayout(strBuf)
					Erase strBuf
				End If
			Next 
			sstType.Visible = True
		End With
		
		Call PaintLine()
		
	End Sub
	
	'-- 폰트 설정
	Private Sub cmdFont_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdFont.Click
		Dim Index As Short = cmdFont.GetIndex(eventSender)
		
		'Cancel을 True로 설정합니다.
		'UPGRADE_WARNING: Visual Basic .NET에서는 CommonDialog CancelError 속성이 지원되지 않습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
        'CommonDialog1.CancelError = True
		On Error GoTo ErrHandler
		
		'Flags 속성을 설정합니다.
		'UPGRADE_WARNING: MSComDlg.CommonDialog 속성 CommonDialog1.Flags이(가) 새로운 동작을 가진 CommonDialog1Font.ShowEffects(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		CommonDialog1Font.ShowEffects = True
		'UPGRADE_ISSUE: cdlCFBoth 상수가 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: MSComDlg.CommonDialog 속성 CommonDialog1.flags이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'CommonDialog1.Flags = MSComDlg.FontsConstants.cdlCFBoth
		
		'폰트 속성을 설정합니다.[Default]
		CommonDialog1Font.Font = VB6.FontChangeName(CommonDialog1Font.Font, "굴림")
		CommonDialog1Font.Font = VB6.FontChangeSize(CommonDialog1Font.Font, 9)
		
		'[글꼴] 대화 상자를 표시합니다.
		CommonDialog1Font.ShowDialog()
		'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
		txtFontName(Index).Text = CommonDialog1Font.Font.Name
		'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
		txtFontSize(Index).Text = CStr(CommonDialog1Font.Font.Size)
		'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
		chkFontBold(Index).CheckState = IIf(CommonDialog1Font.Font.Bold = True, 1, 0)
		'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
		chkFontItalic(Index).CheckState = IIf(CommonDialog1Font.Font.Italic = True, 1, 0)
		'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
		chkFontUnder(Index).CheckState = IIf(CommonDialog1Font.Font.Underline = True, 1, 0)
		
		Exit Sub
		
ErrHandler: 
		'" 사용자가 [취소] 단추를 눌렀습니다.
		Exit Sub
		
	End Sub
	
	'-- 이미지 경로 설정
	Private Sub cmdImage_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdImage.Click
		Dim Index As Short = cmdImage.GetIndex(eventSender)
		
		Dim sFile As String
		sFile = ShowOpen("JPG파일(*.jpg)|*.jpg", My.Application.Info.DirectoryPath & "\" & gImage)
		If sFile <> "" Then
			'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
			txtImageName(Index).Text = sFile
			If Index = 0 Then
				'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
				Didim_SImg.Image = System.Drawing.Image.FromFile(txtImageName(Index).Text)
				'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
				txtImageWSize(Index).Text = CStr(System.Math.Round(VB6.PixelsToTwipsX(Didim_SImg.Width) / CDbl(gScaleCal), 0))
				'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
				txtImageHSize(Index).Text = CStr(System.Math.Round(VB6.PixelsToTwipsY(Didim_SImg.Height) / CDbl(gScaleCal), 0))
				
				'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
				txtImageWSize(Index + 2).Text = txtImageWSize(Index).Text
				'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
				txtImageHSize(Index + 2).Text = txtImageHSize(Index).Text
				
				'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
				txtImageDevide(Index).Focus()
			Else
				'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
				Didim_DImg.Image = System.Drawing.Image.FromFile(txtImageName(Index).Text)
				'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
				txtImageWSize(Index).Text = CStr(System.Math.Round(VB6.PixelsToTwipsX(Didim_DImg.Width) / CDbl(gScaleCal), 0))
				'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
				txtImageHSize(Index).Text = CStr(System.Math.Round(VB6.PixelsToTwipsY(Didim_DImg.Height) / CDbl(gScaleCal), 0))
				
				'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
				txtImageWSize(Index + 2).Text = txtImageWSize(Index).Text
				'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
				txtImageHSize(Index + 2).Text = txtImageHSize(Index).Text
				
				'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
				txtImageDevide(Index).Focus()
			End If
		Else
			'        MsgBox "You pressed cancel"
		End If
		
		
		
		
		'
		'
		'Dim x
		'    'Cancel을 True로 설정합니다.
		'    CommonDialog1.CancelError = True
		'    On Error GoTo ErrHandler
		'
		'    'Flags 속성을 설정합니다.
		'    CommonDialog1.flags = cdlCFEffects Or cdlCFBoth
		'
		'    '경로 속성을 설정합니다.
		'    CommonDialog1.InitDir = App.Path & "\" & gImage
		'
		'    CommonDialog1.Filter = "JPG파일(*.jpg)|*.jpg"
		'
		'    '[파일] 대화 상자를 표시합니다.
		'    CommonDialog1.ShowOpen
		'    txtImageName(Index).Text = CommonDialog1.FileName
		'
		'    If Index = 0 Then
		'        Didim_SImg.Picture = LoadPicture(txtImageName(Index).Text)
		'        txtImageWSize(Index).Text = Round(Didim_SImg.Width / gScaleCal, 0)
		'        txtImageHSize(Index).Text = Round(Didim_SImg.Height / gScaleCal, 0)
		'    Else
		'        Didim_DImg.Picture = LoadPicture(txtImageName(Index).Text)
		'        txtImageWSize(Index).Text = Round(Didim_DImg.Width / gScaleCal, 0)
		'        txtImageHSize(Index).Text = Round(Didim_DImg.Height / gScaleCal, 0)
		'    End If
		'
		'    Exit Sub
		'
		'ErrHandler:
		'  '" 사용자가 [취소] 단추를 눌렀습니다.
		'  Exit Sub
		
	End Sub
	
	'FIXIT: 'obj'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                               FixIT90210ae-R1672-R1B8ZE
	Private Sub MakeSpdSaveList(ByRef obj As Object, ByRef idx As Short)
		
		With spdList
			.MaxRows = .MaxRows + 1
			.Action = FPSpread.ActionConstants.ActionActiveCell
			Select Case idx
				Case 0, 1
					.SetText(1, .MaxRows, .MaxRows - 1) '설정순번
					.SetText(2, .MaxRows, idx) '항목구분
					.SetText(3, .MaxRows, txtTitle.Text) '항목명
					.SetText(4, .MaxRows, txtXpos.Text) 'X1좌표
					.SetText(5, .MaxRows, 0) 'X2좌표
					.SetText(6, .MaxRows, txtYpos.Text) 'Y1좌표
					.SetText(7, .MaxRows, 0) 'Y2좌표
					'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
					.SetText(8, .MaxRows, txtFontName(idx).Text) '폰트명
					'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
					.SetText(9, .MaxRows, txtFontSize(idx).Text) '폰트크기
					'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
					.SetText(10, .MaxRows, IIf(chkFontBold(idx).CheckState = CDbl("0"), "0", "1")) '폰트굵게
					'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
					.SetText(11, .MaxRows, IIf(chkFontUnder(idx).CheckState = CDbl("0"), "0", "1")) '폰트밑줄
					'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
					.SetText(12, .MaxRows, IIf(chkFontItalic(idx).CheckState = CDbl("0"), "0", "1")) '폰트기울게
					.SetText(13, .MaxRows, "0") '폰트회전
					.SetText(14, .MaxRows, "0") '바코드종류
					.SetText(15, .MaxRows, "0") '바코드폭
					.SetText(16, .MaxRows, "0") '바코드회전
					.SetText(17, .MaxRows, "") '이미지경로
					.SetText(18, .MaxRows, "0") '라인회전
					.SetText(19, .MaxRows, "0") '라인두께
					.SetText(20, .MaxRows, "0") '라인폭
					.SetText(21, .MaxRows, IIf(chkPrint.CheckState = CDbl("1"), "0", "1")) '출력여부
					'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
					.SetText(22, .MaxRows, txtContent(idx).Text) '출력값
					.SetText(23, .MaxRows, gScaleCal) 'X좌표 보정값
					.SetText(24, .MaxRows, gScaleCal) 'Y좌표 보정값
					.SetText(25, .MaxRows, txtPaperHSize.Text) '용지높이
					.SetText(26, .MaxRows, txtPaperWSize.Text) '용지폭
					'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
					.SetText(27, .MaxRows, IIf(chkFontItalic(idx).CheckState = CDbl("0"), "0", "1")) '무조건고정
					.SetText(28, .MaxRows, "0") '용지방향
					.SetText(29, .MaxRows, gblCtrlNm) 'Tag
				Case 2
					.SetText(1, .MaxRows, .MaxRows - 1) '설정순번
					.SetText(2, .MaxRows, idx) '항목구분
					.SetText(3, .MaxRows, txtTitle.Text) '항목명
					.SetText(4, .MaxRows, txtXpos.Text) 'X1좌표
					.SetText(5, .MaxRows, txtImageWSize(0).Text) 'X2좌표
					.SetText(6, .MaxRows, txtYpos.Text) 'Y1좌표
					.SetText(7, .MaxRows, txtImageHSize(0).Text) 'Y2좌표
					.SetText(8, .MaxRows, "") '폰트명
					.SetText(9, .MaxRows, "0") '폰트크기
					.SetText(10, .MaxRows, "0") '폰트굵게
					.SetText(11, .MaxRows, "0") '폰트밑줄
					.SetText(12, .MaxRows, "0") '폰트기울게
					.SetText(13, .MaxRows, "0") '폰트회전
					.SetText(14, .MaxRows, "0") '바코드종류
					.SetText(15, .MaxRows, "0") '바코드폭
					.SetText(16, .MaxRows, "0") '바코드회전
					.SetText(17, .MaxRows, txtImageName(0).Text) '이미지경로
					.SetText(18, .MaxRows, "0") '라인회전
					.SetText(19, .MaxRows, "0") '라인두께
					.SetText(20, .MaxRows, "0") '라인폭
					.SetText(21, .MaxRows, IIf(chkPrint.CheckState = CDbl("1"), "0", "1")) '출력여부
					.SetText(22, .MaxRows, "") '출력값
					.SetText(23, .MaxRows, gScaleCal) 'X좌표 보정값
					.SetText(24, .MaxRows, gScaleCal) 'Y좌표 보정값
					.SetText(25, .MaxRows, txtPaperHSize.Text) '용지높이
					.SetText(26, .MaxRows, txtPaperWSize.Text) '용지폭
					.SetText(27, .MaxRows, IIf(chkIStatic.CheckState = CDbl("0"), "0", "1")) '무조건고정
					.SetText(28, .MaxRows, "0") '용지방향
					.SetText(29, .MaxRows, gblCtrlNm) 'Tag
				Case 3
					.SetText(1, .MaxRows, .MaxRows - 1) '설정순번
					.SetText(2, .MaxRows, idx) '항목구분
					.SetText(3, .MaxRows, txtTitle.Text) '항목명
					.SetText(4, .MaxRows, txtXpos.Text) 'X1좌표
					.SetText(5, .MaxRows, txtImageWSize(1).Text) 'X2좌표
					.SetText(6, .MaxRows, txtYpos.Text) 'Y1좌표
					.SetText(7, .MaxRows, txtImageHSize(1).Text) 'Y2좌표
					.SetText(8, .MaxRows, "") '폰트명
					.SetText(9, .MaxRows, "0") '폰트크기
					.SetText(10, .MaxRows, "0") '폰트굵게
					.SetText(11, .MaxRows, "0") '폰트밑줄
					.SetText(12, .MaxRows, "0") '폰트기울게
					.SetText(13, .MaxRows, "0") '폰트회전
					.SetText(14, .MaxRows, "0") '바코드종류
					.SetText(15, .MaxRows, "0") '바코드폭
					.SetText(16, .MaxRows, "0") '바코드회전
					.SetText(17, .MaxRows, txtImageName(1).Text) '이미지경로
					.SetText(18, .MaxRows, "0") '라인회전
					.SetText(19, .MaxRows, "0") '라인두께
					.SetText(20, .MaxRows, "0") '라인폭
					.SetText(21, .MaxRows, IIf(chkPrint.CheckState = CDbl("1"), "0", "1")) '출력여부
					.SetText(22, .MaxRows, "") '출력값
					.SetText(23, .MaxRows, gScaleCal) 'X좌표 보정값
					.SetText(24, .MaxRows, gScaleCal) 'Y좌표 보정값
					.SetText(25, .MaxRows, txtPaperHSize.Text) '용지높이
					.SetText(26, .MaxRows, txtPaperWSize.Text) '용지폭
					.SetText(27, .MaxRows, IIf(chkIStatic.CheckState = CDbl("0"), "0", "1")) '무조건고정
					.SetText(28, .MaxRows, "0") '용지방향
					.SetText(29, .MaxRows, gblCtrlNm) 'Tag
					
				Case 4
					.SetText(1, .MaxRows, .MaxRows - 1) '설정순번
					.SetText(2, .MaxRows, idx) '항목구분
					.SetText(3, .MaxRows, txtTitle.Text) '항목명
					.SetText(4, .MaxRows, txtXpos.Text) 'X1좌표
					.SetText(5, .MaxRows, txtBarWSize.Text) 'X2좌표
					.SetText(6, .MaxRows, txtYpos.Text) 'Y1좌표
					.SetText(7, .MaxRows, txtBarHSize.Text) 'Y2좌표
					.SetText(8, .MaxRows, "") '폰트명
					.SetText(9, .MaxRows, "0") '폰트크기
					.SetText(10, .MaxRows, "0") '폰트굵게
					.SetText(11, .MaxRows, "0") '폰트밑줄
					.SetText(12, .MaxRows, "0") '폰트기울게
					.SetText(13, .MaxRows, "0") '폰트회전
					.SetText(14, .MaxRows, cboBarType.SelectedIndex) '바코드종류
					.SetText(15, .MaxRows, "0") 'txtBarDevide.Text                           '바코드폭
					.SetText(16, .MaxRows, IIf(chkBarRotate.CheckState = CDbl("0"), 0, 2)) '바코드회전
					.SetText(17, .MaxRows, "") '이미지경로
					.SetText(18, .MaxRows, "0") '라인회전
					.SetText(19, .MaxRows, "0") '라인두께
					.SetText(20, .MaxRows, "0") '라인폭
					.SetText(21, .MaxRows, IIf(chkPrint.CheckState = CDbl("1"), "0", "1")) '출력여부
					'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
					.SetText(22, .MaxRows, Trim(txtBarData.Text)) '출력값
					.SetText(23, .MaxRows, gScaleCal) 'X좌표 보정값
					.SetText(24, .MaxRows, gScaleCal) 'Y좌표 보정값
					.SetText(25, .MaxRows, txtPaperHSize.Text) '용지높이
					.SetText(26, .MaxRows, txtPaperWSize.Text) '용지폭
					.SetText(27, .MaxRows, IIf(chkIStatic.CheckState = CDbl("0"), "0", "1")) '무조건고정
					.SetText(28, .MaxRows, "0") '용지방향
					.SetText(29, .MaxRows, gblCtrlNm) 'Tag
					
				Case 5
					.SetText(1, .MaxRows, .MaxRows - 1) '설정순번
					.SetText(2, .MaxRows, idx) '항목구분
					.SetText(3, .MaxRows, txtTitle.Text) '항목명
					If chkLineRotate.CheckState = CDbl("0") Then
						.SetText(4, .MaxRows, txtXpos.Text) 'X1좌표
						.SetText(5, .MaxRows, txtLineWSize.Text) 'X2좌표
						.SetText(6, .MaxRows, txtYpos.Text) 'Y1좌표
						.SetText(7, .MaxRows, txtYpos.Text) 'Y2좌표
					Else
						.SetText(4, .MaxRows, txtXpos.Text) 'X1좌표
						.SetText(5, .MaxRows, txtXpos.Text) 'X2좌표
						.SetText(6, .MaxRows, txtYpos.Text) 'Y1좌표
						.SetText(7, .MaxRows, txtLineWSize.Text) 'Y2좌표
					End If
					.SetText(8, .MaxRows, "") '폰트명
					.SetText(9, .MaxRows, "1") '폰트크기
					.SetText(10, .MaxRows, "0") '폰트굵게
					.SetText(11, .MaxRows, "0") '폰트밑줄
					.SetText(12, .MaxRows, "0") '폰트기울게
					.SetText(13, .MaxRows, "0") '폰트회전
					.SetText(14, .MaxRows, "0") '바코드종류
					.SetText(15, .MaxRows, "0") '바코드폭
					.SetText(16, .MaxRows, "0") '바코드회전
					.SetText(17, .MaxRows, "") '이미지경로
					.SetText(18, .MaxRows, IIf(chkLineRotate.CheckState = CDbl("0"), "0", "1")) '라인회전
					.SetText(19, .MaxRows, txtLineHSize.Text) '라인두께
					.SetText(20, .MaxRows, txtLineWSize.Text) '라인폭
					.SetText(21, .MaxRows, IIf(chkPrint.CheckState = CDbl("1"), "0", "1")) '출력여부
					.SetText(22, .MaxRows, "") '출력값
					.SetText(23, .MaxRows, gScaleCal) 'X좌표 보정값
					.SetText(24, .MaxRows, gScaleCal) 'Y좌표 보정값
					.SetText(25, .MaxRows, txtPaperHSize.Text) '용지높이
					.SetText(26, .MaxRows, txtPaperWSize.Text) '용지폭
					.SetText(27, .MaxRows, IIf(chkIStatic.CheckState = CDbl("0"), "0", "1")) '무조건고정
					.SetText(28, .MaxRows, "0") '용지방향
					.SetText(29, .MaxRows, gblCtrlNm) 'Tag
					
			End Select
			
			'        .ColWidth(-1) = 5
		End With
		
	End Sub
	
	' 오프젝트를 생성시킨다.
	Private Function objMake() As String
		'FIXIT: 'obj'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                               FixIT90210ae-R1672-R1B8ZE
		Dim obj As Object
		Dim ClsEventObject As ClassEventObject
		
		ClsEventObject = New ClassEventObject
		
		objMake = "0"
		
		Select Case sstType.SelectedIndex
			Case 0 'Static Label
				obj = ClsEventObject.CreateObject_Renamed(Me, ClsEventMonitor, ClassEventMonitor.EventObjectID.EventObjectSLabel, txtTag.Text)
				If Not obj Is Nothing Then
					'UPGRADE_WARNING: obj.Tag 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Tag = txtTitle.Text
					'UPGRADE_WARNING: obj.AutoSize 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.AutoSize = True
					'UPGRADE_WARNING: obj.BackColor 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
					'UPGRADE_WARNING: obj.Font 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Font = txtFontName(sstType.SelectedIndex).Text
					'UPGRADE_WARNING: obj.FontSize 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.FontSize = System.Math.Round(CDbl(txtFontSize(sstType.SelectedIndex).Text) * CDbl(gDevide), 0)
					'UPGRADE_WARNING: obj.FontBold 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.FontBold = IIf(chkFontBold(sstType.SelectedIndex).CheckState = 1, True, False)
					'UPGRADE_WARNING: obj.FontItalic 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.FontItalic = IIf(chkFontItalic(sstType.SelectedIndex).CheckState = 1, True, False)
					'UPGRADE_WARNING: obj.FontUnderline 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.FontUnderline = IIf(chkFontUnder(sstType.SelectedIndex).CheckState = 1, True, False)
					'UPGRADE_WARNING: obj.Top 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Top = System.Math.Round(CDbl(txtYpos.Text) * CDbl(gDevide), 0)
					'UPGRADE_WARNING: obj.Left 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Left = System.Math.Round(CDbl(txtXpos.Text) * CDbl(gDevide), 0)
					'UPGRADE_WARNING: obj.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Caption = txtContent(sstType.SelectedIndex).Text
					'UPGRADE_WARNING: obj.DataMember 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.DataMember = chkTStatic.CheckState '-- 무조건고정
					'UPGRADE_WARNING: obj.DataField 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.DataField = IIf(chkPrint.CheckState = CDbl("1"), "0", "1") '-- 출력안함
					'UPGRADE_WARNING: obj.MousePointer 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.MousePointer = 5
					
				Else
					MsgBox("동일한 항목명은 사용할 수 없습니다.", MsgBoxStyle.Information, Me.Text)
					'UPGRADE_NOTE: ClsEventObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					ClsEventObject = Nothing
					Exit Function
					'            Set ClsEventObject = Nothing
					'            If MsgBox("동일한 항목명은 사용할 수 없습니다." & vbNewLine & "종료하시겠습니까?", vbYesNo + vbCritical, Me.Caption) = vbYes Then
					'                objMake = txtTag.Text & "_EDIT"
					'                Exit Function
					'            End If
				End If
			Case 1 'Dynamic Label
				obj = ClsEventObject.CreateObject_Renamed(Me, ClsEventMonitor, ClassEventMonitor.EventObjectID.EventObjectDLabel, txtTag.Text)
				If Not obj Is Nothing Then
					'UPGRADE_WARNING: obj.Tag 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Tag = txtTitle.Text
					'UPGRADE_WARNING: obj.AutoSize 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.AutoSize = True
					'UPGRADE_WARNING: obj.BackColor 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
					'UPGRADE_WARNING: obj.Font 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Font = txtFontName(sstType.SelectedIndex).Text
					'UPGRADE_WARNING: obj.FontSize 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.FontSize = System.Math.Round(CDbl(txtFontSize(sstType.SelectedIndex).Text) * CDbl(gDevide), 0)
					'UPGRADE_WARNING: obj.FontBold 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.FontBold = IIf(chkFontBold(sstType.SelectedIndex).CheckState = 1, True, False)
					'UPGRADE_WARNING: obj.FontItalic 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.FontItalic = IIf(chkFontItalic(sstType.SelectedIndex).CheckState = 1, True, False)
					'UPGRADE_WARNING: obj.FontUnderline 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.FontUnderline = IIf(chkFontUnder(sstType.SelectedIndex).CheckState = 1, True, False)
					'UPGRADE_WARNING: obj.Top 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Top = System.Math.Round(CDbl(txtYpos.Text) * CDbl(gDevide), 0)
					'UPGRADE_WARNING: obj.Left 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Left = System.Math.Round(CDbl(txtXpos.Text) * CDbl(gDevide), 0)
					'UPGRADE_WARNING: obj.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Caption = txtContent(sstType.SelectedIndex).Text
					'UPGRADE_WARNING: obj.DataMember 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.DataMember = IIf(chkPrint.CheckState = CDbl("1"), "0", "1") '-- 출력안함
					'UPGRADE_WARNING: obj.MousePointer 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.MousePointer = 5
				Else
					MsgBox("동일한 항목명은 사용할 수 없습니다.", MsgBoxStyle.Information, Me.Text)
					'UPGRADE_NOTE: ClsEventObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					ClsEventObject = Nothing
					Exit Function
					'            Set ClsEventObject = Nothing
					'            If MsgBox(txtTag.Text & " 항목명은 사용할 수 없습니다." & vbNewLine & "항목명을 변경하시겠습니까?", vbYesNo + vbInformation, Me.Caption) = vbYes Then
					'                objMake = txtTag.Text & "_EDIT"
					'                Exit Function
					'            End If
				End If
			Case 2 'Static Image
				obj = ClsEventObject.CreateObject_Renamed(Me, ClsEventMonitor, ClassEventMonitor.EventObjectID.EventObjectSImage, txtTag.Text)
				If Not obj Is Nothing Then
					'UPGRADE_WARNING: Dir에 새 동작이 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					If Dir(txtImageName(0).Text) = "" Then
						'UPGRADE_WARNING: obj.Picture 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Picture = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\" & gImage & "noimage.bmp")
					Else
						'UPGRADE_WARNING: obj.Picture 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Picture = System.Drawing.Image.FromFile(txtImageName(0).Text)
					End If
					'UPGRADE_WARNING: obj.Tag 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Tag = txtTitle.Text
					'UPGRADE_WARNING: obj.DataMember 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.DataMember = txtImageName(0).Text '-- 이미지경로
					'UPGRADE_WARNING: obj.Stretch 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Stretch = True
					'UPGRADE_WARNING: obj.Width 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Width = System.Math.Round(CDbl(txtImageWSize(0).Text) * CDbl(gDevide), 0)
					'UPGRADE_WARNING: obj.Height 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Height = System.Math.Round(CDbl(txtImageHSize(0).Text) * CDbl(gDevide), 0)
					'UPGRADE_WARNING: obj.Top 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Top = System.Math.Round(CDbl(txtYpos.Text) * CDbl(gDevide), 0)
					'UPGRADE_WARNING: obj.Left 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Left = System.Math.Round(CDbl(txtXpos.Text) * CDbl(gDevide), 0)
					'UPGRADE_WARNING: obj.ToolTipText 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.ToolTipText = CStr(chkIStatic.CheckState) '-- 무조건고정
					'UPGRADE_WARNING: obj.DataField 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.DataField = IIf(chkPrint.CheckState = CDbl("1"), "0", "1") '-- 출력안함
					'UPGRADE_WARNING: obj.MousePointer 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.MousePointer = 5
				Else
					MsgBox("동일한 항목명은 사용할 수 없습니다.", MsgBoxStyle.Information, Me.Text)
					'UPGRADE_NOTE: ClsEventObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					ClsEventObject = Nothing
					Exit Function
					'            Set ClsEventObject = Nothing
					'            If MsgBox(txtTag.Text & " 항목명은 사용할 수 없습니다." & vbNewLine & "항목명을 변경하시겠습니까?", vbYesNo + vbInformation, Me.Caption) = vbYes Then
					'                objMake = txtTag.Text & "_EDIT"
					'                Exit Function
					'            End If
					
				End If
			Case 3 'Dynamic Image
				obj = ClsEventObject.CreateObject_Renamed(Me, ClsEventMonitor, ClassEventMonitor.EventObjectID.EventObjectDImage, txtTag.Text)
				If Not obj Is Nothing Then
					'UPGRADE_WARNING: Dir에 새 동작이 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					If Dir(txtImageName(1).Text) = "" Then
						'UPGRADE_WARNING: obj.Picture 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Picture = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\" & gImage & "noimage.bmp")
					Else
						'UPGRADE_WARNING: obj.Picture 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Picture = System.Drawing.Image.FromFile(txtImageName(1).Text)
					End If
					'UPGRADE_WARNING: obj.Tag 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Tag = txtTitle.Text
					'UPGRADE_WARNING: obj.DataMember 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.DataMember = txtImageName(1).Text '-- 이미지경로
					'UPGRADE_WARNING: obj.Stretch 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Stretch = True
					'UPGRADE_WARNING: obj.Width 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Width = System.Math.Round(CDbl(txtImageWSize(1).Text) * CDbl(gDevide), 0)
					'UPGRADE_WARNING: obj.Height 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Height = System.Math.Round(CDbl(txtImageHSize(1).Text) * CDbl(gDevide), 0)
					'UPGRADE_WARNING: obj.Top 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Top = System.Math.Round(CDbl(txtYpos.Text) * CDbl(gDevide), 0)
					'UPGRADE_WARNING: obj.Left 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Left = System.Math.Round(CDbl(txtXpos.Text) * CDbl(gDevide), 0)
					'UPGRADE_WARNING: obj.DataField 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.DataField = IIf(chkPrint.CheckState = CDbl("1"), "0", "1") '-- 출력안함
					'UPGRADE_WARNING: obj.MousePointer 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.MousePointer = 5
				Else
					MsgBox("동일한 항목명은 사용할 수 없습니다.", MsgBoxStyle.Information, Me.Text)
					'UPGRADE_NOTE: ClsEventObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					ClsEventObject = Nothing
					Exit Function
					'            Set ClsEventObject = Nothing
					'            If MsgBox(txtTag.Text & " 항목명은 사용할 수 없습니다." & vbNewLine & "항목명을 변경하시겠습니까?", vbYesNo + vbInformation, Me.Caption) = vbYes Then
					'                objMake = txtTag.Text & "_EDIT"
					'                Exit Function
					'            End If
					
				End If
				
			Case 4 'Barcode
				obj = ClsEventObject.CreateObject_Renamed(Me, ClsEventMonitor, ClassEventMonitor.EventObjectID.EventObjectBarcode, txtTag.Text)
				If Not obj Is Nothing Then
					'UPGRADE_WARNING: obj.Tag 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Tag = txtTitle.Text
					'UPGRADE_WARNING: obj.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Caption = txtBarData.Text
					'UPGRADE_WARNING: obj.Style 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Style = cboBarType.SelectedIndex
					'UPGRADE_WARNING: obj.Alignment 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Alignment = BarcodLib.AlignmentConstants.bcALeft
					'UPGRADE_WARNING: obj.BarWidth 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.BarWidth = 0
					'UPGRADE_WARNING: obj.Width 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Width = System.Math.Round(CDbl(txtBarWSize.Text) * CDbl(gDevide), 0)
					'UPGRADE_WARNING: obj.Height 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Height = System.Math.Round(CDbl(txtBarHSize.Text) * CDbl(gDevide), 0)
					'UPGRADE_WARNING: obj.Top 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Top = System.Math.Round(CDbl(txtYpos.Text) * CDbl(gDevide), 0)
					'UPGRADE_WARNING: obj.Left 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Left = System.Math.Round(CDbl(txtXpos.Text) * CDbl(gDevide), 0)
					'UPGRADE_WARNING: obj.Direction 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Direction = IIf(chkBarRotate.CheckState = CDbl("0"), 0, 2)
					'UPGRADE_WARNING: obj.Visible 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Visible = False
					'            obj.Visible = True
					
					'UPGRADE_WARNING: obj.Container 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Container = Picture1
					m_ColCommandButton.Add(ClsEventObject)
					'UPGRADE_NOTE: ClsEventObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					ClsEventObject = Nothing
					
					'== 바코드를 이미지 형태로 올리기 ===================================================================
					If intMode = 0 Then '==== Mode Set [0:로드,1:적용,2:이동,3:생성]
						If strBarImgName = "" Then
							'strBarImgName = txtTag.Text & "_IMG1"
							strBarImgName = txtTag.Text & "_IMG"
						Else
							'FIXIT: 'Right' 함수를 'Right$' 함수로 바꾸십시오.                                                    FixIT90210ae-R9757-R1B8ZE
							'FIXIT: 'Mid' 함수를 'Mid$' 함수로 바꾸십시오.                                                        FixIT90210ae-R9757-R1B8ZE
							strBarImgName = Mid(strBarImgName, 1, Len(strBarImgName) - 1) & CDbl(VB.Right(strBarImgName, 1)) + 1
						End If
					End If
					
					ClsEventObject = New ClassEventObject
					obj = ClsEventObject.CreateObject_Renamed(Me, ClsEventMonitor, ClassEventMonitor.EventObjectID.EventObjectBImage, strBarImgName)
					If Not obj Is Nothing Then
						'UPGRADE_WARNING: obj.Tag 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Tag = txtTitle.Text
						'UPGRADE_WARNING: obj.Stretch 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Stretch = True
						If chkBarRotate.CheckState = CDbl("0") Then
							'UPGRADE_WARNING: obj.Picture 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							obj.Picture = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\" & gImage & "\barcode.bmp")
							'UPGRADE_WARNING: obj.DataMember 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							obj.DataMember = My.Application.Info.DirectoryPath & "\" & gImage & "\barcode.bmp" '-- 이미지 경로
							'UPGRADE_WARNING: obj.Width 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							obj.Width = System.Math.Round(CDbl(txtBarWSize.Text) * CDbl(gDevide), 0)
							'UPGRADE_WARNING: obj.Height 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							obj.Height = System.Math.Round(CDbl(txtBarHSize.Text) * CDbl(gDevide), 0)
						Else
							'UPGRADE_WARNING: obj.Picture 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							obj.Picture = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\" & gImage & "\barcode90.bmp")
							'UPGRADE_WARNING: obj.DataMember 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							obj.DataMember = My.Application.Info.DirectoryPath & "\" & gImage & "\barcode90.bmp" '-- 이미지 경로
							'UPGRADE_WARNING: obj.Width 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							obj.Width = System.Math.Round(CDbl(txtBarHSize.Text) * CDbl(gDevide), 0)
							'UPGRADE_WARNING: obj.Height 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							obj.Height = System.Math.Round(CDbl(txtBarWSize.Text) * CDbl(gDevide), 0)
						End If
						
						
						'                obj.Width = Round(txtBarWSize.Text * gDevide, 0)
						'                obj.Height = Round(txtBarHSize.Text * gDevide, 0)
						'UPGRADE_WARNING: obj.Top 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Top = System.Math.Round(CDbl(txtYpos.Text) * CDbl(gDevide), 0)
						'UPGRADE_WARNING: obj.Left 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Left = System.Math.Round(CDbl(txtXpos.Text) * CDbl(gDevide), 0)
						'UPGRADE_WARNING: obj.ToolTipText 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.ToolTipText = CStr(cboBarType.SelectedIndex) '-- 바코드 타입
						'UPGRADE_WARNING: obj.DataField 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.DataField = IIf(chkPrint.CheckState = CDbl("1"), "0", "1") '-- 출력안함
						'UPGRADE_WARNING: obj.MousePointer 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.MousePointer = 5
					Else
						MsgBox("동일한 항목명은 사용할 수 없습니다.[바코드 생성 오류]", MsgBoxStyle.Information, Me.Text)
						'UPGRADE_NOTE: ClsEventObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						ClsEventObject = Nothing
						Exit Function
					End If
					'== 바코드를 이미지 형태로 올리기 ===================================================================
				Else
					MsgBox("동일한 항목명은 사용할 수 없습니다.", MsgBoxStyle.Information, Me.Text)
					'UPGRADE_NOTE: ClsEventObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					ClsEventObject = Nothing
					Exit Function
				End If
			Case 5 'Line
				obj = ClsEventObject.CreateObject_Renamed(Me, ClsEventMonitor, ClassEventMonitor.EventObjectID.EventObjectLImage, txtTag.Text)
				If Not obj Is Nothing Then
					'UPGRADE_WARNING: obj.Tag 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Tag = txtTitle.Text
					If chkLineRotate.CheckState = 0 Then
						'UPGRADE_WARNING: obj.Picture 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Picture = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\" & gImage & "wline.jpg")
						'UPGRADE_WARNING: obj.Stretch 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Stretch = True
						'UPGRADE_WARNING: obj.Width 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Width = System.Math.Round(CDbl(txtLineWSize.Text) * CDbl(gDevide), 0)
						'UPGRADE_WARNING: obj.Height 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Height = System.Math.Round(CDbl(txtLineHSize.Text) * CDbl(gDevide), 0)
						'UPGRADE_WARNING: obj.Top 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Top = System.Math.Round(CDbl(txtYpos.Text) * CDbl(gDevide), 0)
						'UPGRADE_WARNING: obj.Left 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Left = System.Math.Round(CDbl(txtXpos.Text) * CDbl(gDevide), 0)
						'UPGRADE_WARNING: obj.ToolTipText 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.ToolTipText = IIf(chkPrint.CheckState = CDbl("1"), "0", "1") '-- 출력안함
						'UPGRADE_WARNING: obj.DataMember 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.DataMember = "0" '-- Rotate
						'UPGRADE_WARNING: obj.MousePointer 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.MousePointer = 5
					Else
						'UPGRADE_WARNING: obj.Picture 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Picture = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\" & gImage & "hline.jpg")
						'UPGRADE_WARNING: obj.Stretch 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Stretch = True
						'UPGRADE_WARNING: obj.Width 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Width = System.Math.Round(CDbl(txtLineHSize.Text) * CDbl(gDevide), 0)
						'UPGRADE_WARNING: obj.Height 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Height = System.Math.Round(CDbl(txtLineWSize.Text) * CDbl(gDevide), 0)
						'UPGRADE_WARNING: obj.Top 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Top = System.Math.Round(CDbl(txtYpos.Text) * CDbl(gDevide), 0)
						'UPGRADE_WARNING: obj.Left 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Left = System.Math.Round(CDbl(txtXpos.Text) * CDbl(gDevide), 0)
						'UPGRADE_WARNING: obj.ToolTipText 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.ToolTipText = IIf(chkPrint.CheckState = CDbl("1"), "0", "1") '-- 출력안함
						'UPGRADE_WARNING: obj.DataMember 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.DataMember = "1" '-- Rotate
						'UPGRADE_WARNING: obj.MousePointer 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.MousePointer = 5
					End If
				Else
					MsgBox("동일한 항목명은 사용할 수 없습니다.", MsgBoxStyle.Information, Me.Text)
					'UPGRADE_NOTE: ClsEventObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					ClsEventObject = Nothing
					Exit Function
					'            Set ClsEventObject = Nothing
					'            If MsgBox(txtTag.Text & " 항목명은 사용할 수 없습니다." & vbNewLine & "항목명을 변경하시겠습니까?", vbYesNo + vbInformation, Me.Caption) = vbYes Then
					'                objMake = txtTag.Text & "_EDIT"
					'                Exit Function
					'            End If
				End If
		End Select
		
		'UPGRADE_WARNING: obj.Visible 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj.Visible = True
		'UPGRADE_WARNING: obj.Container 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj.Container = Picture1
		m_ColCommandButton.Add(ClsEventObject)
		'UPGRADE_NOTE: ClsEventObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		ClsEventObject = Nothing
		
	End Function
	
	'FIXIT: 'BarObj'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                            FixIT90210ae-R1672-R1B8ZE
	Private Sub MakeBarImage(ByVal BarObj As Object)
		
		'UPGRADE_WARNING: BarObj.Height 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Picture2.Height = VB6.TwipsToPixelsY(BarObj.Height)
		'UPGRADE_WARNING: BarObj.Width 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Picture2.Width = VB6.TwipsToPixelsX(BarObj.Width)
		'UPGRADE_ISSUE: vbTwips 상수가 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
        'Barcod1.PrinterScaleMode = vbTwips 'Form1.ScaleMode
		'UPGRADE_WARNING: BarObj.Width 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Barcod1.PrinterWidth = BarObj.Width
		'UPGRADE_WARNING: BarObj.Height 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Barcod1.PrinterHeight = BarObj.Height
		Barcod1.PrinterTop = 0
		Barcod1.PrinterLeft = 0
		'UPGRADE_ISSUE: PictureBox 속성 Picture2.hdc이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Barcod1.PrinterHDC = Picture2.hdc
		Picture2.Refresh()
		'FIXIT: Clipboard 개체는 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.                      FixIT90210ae-R6194-H1984
		My.Computer.Clipboard.Clear()
		'FIXIT: Clipboard 개체는 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.                      FixIT90210ae-R6194-H1984
		'FIXIT: Picture2.Image property 은(는) Visual Basic .NET에서 해당되는 항목이 없으므로 업그레이드되지 않습니다.       FixIT90210ae-R7593-R67265
		'UPGRADE_ISSUE: PictureBox 속성 Picture2.Image이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		My.Computer.Clipboard.SetImage(Picture2.Image)
		
		'    SavePicture Picture2.Image, "C:\TEST.BMP"
		'FIXIT: Picture2.Image property 은(는) Visual Basic .NET에서 해당되는 항목이 없으므로 업그레이드되지 않습니다.       FixIT90210ae-R7593-R67265
		'UPGRADE_WARNING: SavePicture이(가) System.Drawing.Image.Save(으)로 업그레이드되어 새 동작을 가집니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Picture2.Image.Save("C:\TEST.BMP")
		
	End Sub
	
	Private Function findSameCtrlNm(ByRef strIdx As String, ByRef strTitle As String) As Boolean
		Dim i As Short
		Dim strCtrlIdx As String
		Dim strCtrlNm As String
		
		findSameCtrlNm = False
		With spdList
			For i = 1 To .MaxRows
				.Row = i
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				.Col = 2 : strCtrlIdx = Trim(.Text)
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				.Col = 3 : strCtrlNm = Trim(.Text)
				If strIdx = strCtrlIdx And strTitle = strCtrlNm Then
					findSameCtrlNm = True
					Exit For
				End If
			Next 
		End With
		
	End Function
	
	Private Sub objNewMake()
		'FIXIT: 'obj'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                               FixIT90210ae-R1672-R1B8ZE
		Dim obj As Object
		Dim i As Short
		Dim ClsEventObject As ClassEventObject
		
		'-- 유효성 검사 [항목명]
		'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
		If Trim(txtTitle.Text) = "" Then
			MsgBox("항목명을 입력하세요.", MsgBoxStyle.Information, Me.Text)
			txtTitle.Focus()
			Exit Sub
		End If
		'-- 유효성 검사 [X 좌표명]
		'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
		If Trim(txtXpos.Text) = "" Then
			MsgBox("X좌표를 입력하세요.", MsgBoxStyle.Information, Me.Text)
			txtXpos.Focus()
			Exit Sub
		End If
		'-- 유효성 검사 [X 좌표]
		'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
		If Not IsNumeric(Trim(txtXpos.Text)) Then
			MsgBox("X좌표는 숫자만 입력이 가능합니다.", MsgBoxStyle.OKOnly + MsgBoxStyle.Information, Me.Text)
			txtXpos.Focus()
			Exit Sub
		End If
		'-- 유효성 검사 [Y 좌표명]
		'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
		If Trim(txtYpos.Text) = "" Then
			MsgBox("Y좌표를 입력하세요.", MsgBoxStyle.Information, Me.Text)
			txtYpos.Focus()
			Exit Sub
		End If
		'-- 유효성 검사 [Y 좌표]
		'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
		If Not IsNumeric(Trim(txtYpos.Text)) Then
			MsgBox("Y좌표는 숫자만 입력이 가능합니다.", MsgBoxStyle.OKOnly + MsgBoxStyle.Information, Me.Text)
			txtYpos.Focus()
			Exit Sub
		End If
		
		Select Case sstType.SelectedIndex
			Case 0 '## Static Label ##
				'-- 유효성 검사 [폰트명]
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				If Trim(txtFontName(0).Text) = "" Or Trim(txtFontSize(0).Text) = "" Then
					MsgBox("Font를 선택하세요.", MsgBoxStyle.Information, Me.Text)
					Call cmdFont_Click(cmdFont.Item(0), New System.EventArgs())
					Exit Sub
				End If
				'-- 유효성 검사 [폰트사이즈]
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				If Not IsNumeric(Trim(txtFontSize(0).Text)) Then
					MsgBox("숫자만 입력이 가능합니다.", MsgBoxStyle.OKOnly + MsgBoxStyle.Information, Me.Text)
					txtFontSize(0).Focus()
					Exit Sub
				End If
				'-- 유효성 검사 [텍스트]
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				If Trim(txtContent(0).Text) = "" Then
					MsgBox("Text를 입력하세요.", MsgBoxStyle.Information, Me.Text)
					txtContent(0).Focus()
					Exit Sub
				End If
				
				'-- 동일명칭 체크
				If findSameCtrlNm(CStr(sstType.SelectedIndex), (txtTitle.Text)) Then
					MsgBox("동일한 항목명은 사용할 수 없습니다.", MsgBoxStyle.Information, Me.Text)
					Exit Sub
				End If
				
				'-- Static Label 개체만들기
				gblCtrlIdx = gblCtrlIdx + 1
				gblCtrlNm = "Control_" & gblCtrlIdx
				
				ClsEventObject = New ClassEventObject
				'Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectSLabel, txtTitle.Text)
				obj = ClsEventObject.CreateObject_Renamed(Me, ClsEventMonitor, ClassEventMonitor.EventObjectID.EventObjectSLabel, gblCtrlNm)
				If Not obj Is Nothing Then
					'UPGRADE_WARNING: obj.Tag 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Tag = txtTitle.Text
					'UPGRADE_WARNING: obj.AutoSize 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.AutoSize = True
					'UPGRADE_WARNING: obj.BackColor 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
					'UPGRADE_WARNING: obj.Font 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Font = txtFontName(sstType.SelectedIndex).Text
					'UPGRADE_WARNING: obj.FontSize 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.FontSize = CDbl(txtFontSize(sstType.SelectedIndex).Text) * CDbl(gDevide)
					'UPGRADE_WARNING: obj.FontBold 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.FontBold = IIf(chkFontBold(sstType.SelectedIndex).CheckState = 1, True, False)
					'UPGRADE_WARNING: obj.FontItalic 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.FontItalic = IIf(chkFontItalic(sstType.SelectedIndex).CheckState = 1, True, False)
					'UPGRADE_WARNING: obj.FontUnderline 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.FontUnderline = IIf(chkFontUnder(sstType.SelectedIndex).CheckState = 1, True, False)
					'UPGRADE_WARNING: obj.Top 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Top = CDbl(txtYpos.Text) * CDbl(gDevide)
					'UPGRADE_WARNING: obj.Left 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Left = CDbl(txtXpos.Text) * CDbl(gDevide)
					'UPGRADE_WARNING: obj.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Caption = txtContent(sstType.SelectedIndex).Text
					'UPGRADE_WARNING: obj.DataMember 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.DataMember = chkTStatic.CheckState '-- 무조건고정
					'UPGRADE_WARNING: obj.DataField 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.DataField = IIf(chkPrint.CheckState = CDbl("1"), "0", "1") '-- 출력안함
					'UPGRADE_WARNING: obj.MousePointer 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.MousePointer = 5
					
					'obj======그리는곳
					'X , Y====좌표
					'Txt======글자
					'TxtGag===글자의 기울기
					'H========글자의 높이(1에 대한 배율)
					'W========글자의 너비(1에 대한 배율)
					'LineSpace ====줄간격(1에 대한 배율)
					
					'                Call RotateControl(obj, 90)
					
					'                If optSTRotate(0).Value = True Then
					'                    Call FontStuff(Picture1, obj.Top, obj.Left, obj.Caption, 0, 1, 1, 1)
					'
					'                ElseIf optSTRotate(1).Value = True Then
					'                    Call FontStuff(Picture1, obj.Top, obj.Left, obj.Caption, 90, 1, 1, 1)
					'                ElseIf optSTRotate(2).Value = True Then
					'                    Call FontStuff(Picture1, obj.Top, obj.Left, obj.Caption, 180, 1, 1, 1)
					'                Else
					'                    Call FontStuff(Picture1, obj.Top, obj.Left, obj.Caption, 270, 1, 1, 1)
					'                End If
					
					
					Call MakeSpdSaveList(obj, (sstType.SelectedIndex))
				Else
					If gblCtrlIdx > 0 Then gblCtrlIdx = gblCtrlIdx - 1
					MsgBox("동일한 항목명은 사용할 수 없습니다.", MsgBoxStyle.Information, Me.Text)
					'UPGRADE_NOTE: ClsEventObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					ClsEventObject = Nothing
					Exit Sub
				End If
				
			Case 1 '## Dynamic Label ##
				'-- 유효성 검사 [폰트명]
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				If Trim(txtFontName(1).Text) = "" Or Trim(txtFontSize(1).Text) = "" Then
					MsgBox("Font를 선택하세요.", MsgBoxStyle.Information, Me.Text)
					Call cmdFont_Click(cmdFont.Item(1), New System.EventArgs())
					Exit Sub
				End If
				'-- 유효성 검사 [폰트사이즈]
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				If Not IsNumeric(Trim(txtFontSize(1).Text)) Then
					MsgBox("숫자만 입력이 가능합니다.", MsgBoxStyle.OKOnly + MsgBoxStyle.Information, Me.Text)
					txtFontSize(1).Focus()
					Exit Sub
				End If
				'-- 유효성 검사 [텍스트]
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				If Trim(txtContent(1).Text) = "" Then
					MsgBox("Text를 입력하세요.", MsgBoxStyle.Information, Me.Text)
					txtContent(1).Focus()
					Exit Sub
				End If
				
				'-- 동일명칭 체크
				If findSameCtrlNm(CStr(sstType.SelectedIndex), (txtTitle.Text)) Then
					MsgBox("동일한 항목명은 사용할 수 없습니다.", MsgBoxStyle.Information, Me.Text)
					Exit Sub
				End If
				
				'-- Dynamic Label 개체만들기
				gblCtrlIdx = gblCtrlIdx + 1
				gblCtrlNm = "Control_" & gblCtrlIdx
				
				ClsEventObject = New ClassEventObject
				'Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectDLabel, txtTitle.Text)
				obj = ClsEventObject.CreateObject_Renamed(Me, ClsEventMonitor, ClassEventMonitor.EventObjectID.EventObjectDLabel, gblCtrlNm)
				If Not obj Is Nothing Then
					'UPGRADE_WARNING: obj.Tag 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Tag = txtTitle.Text
					'UPGRADE_WARNING: obj.AutoSize 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.AutoSize = True
					'UPGRADE_WARNING: obj.BackColor 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
					'UPGRADE_WARNING: obj.Font 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Font = txtFontName(sstType.SelectedIndex).Text
					'UPGRADE_WARNING: obj.FontSize 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.FontSize = CDbl(txtFontSize(sstType.SelectedIndex).Text) * CDbl(gDevide)
					'UPGRADE_WARNING: obj.FontBold 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.FontBold = IIf(chkFontBold(sstType.SelectedIndex).CheckState = 1, True, False)
					'UPGRADE_WARNING: obj.FontItalic 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.FontItalic = IIf(chkFontItalic(sstType.SelectedIndex).CheckState = 1, True, False)
					'UPGRADE_WARNING: obj.FontUnderline 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.FontUnderline = IIf(chkFontUnder(sstType.SelectedIndex).CheckState = 1, True, False)
					'UPGRADE_WARNING: obj.Top 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Top = CDbl(txtYpos.Text) * CDbl(gDevide)
					'UPGRADE_WARNING: obj.Left 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Left = CDbl(txtXpos.Text) * CDbl(gDevide)
					'UPGRADE_WARNING: obj.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Caption = txtContent(sstType.SelectedIndex).Text
					'UPGRADE_WARNING: obj.DataMember 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.DataMember = IIf(chkPrint.CheckState = CDbl("1"), "0", "1") '-- 출력안함
					'UPGRADE_WARNING: obj.MousePointer 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.MousePointer = 5
					Call MakeSpdSaveList(obj, (sstType.SelectedIndex))
				Else
					If gblCtrlIdx > 0 Then gblCtrlIdx = gblCtrlIdx - 1
					MsgBox("동일한 항목명은 사용할 수 없습니다.", MsgBoxStyle.Information, Me.Text)
					'UPGRADE_NOTE: ClsEventObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					ClsEventObject = Nothing
					Exit Sub
				End If
				
			Case 2 '## Static Image ##
				'-- 유효성 검사 [이미지명]
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				If Trim(txtImageName(0).Text) = "" Then
					MsgBox("이미지를 선택하세요.", MsgBoxStyle.Information, Me.Text)
					Call cmdImage_Click(cmdImage.Item(0), New System.EventArgs())
					Exit Sub
				End If
				'-- 유효성 검사 [가로Size]
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				If Trim(txtImageWSize(0).Text) = "" Then
					MsgBox("가로Size를 입력하세요.", MsgBoxStyle.Information, Me.Text)
					txtImageWSize(0).Focus()
					Exit Sub
				End If
				'-- 유효성 검사 [가로Size]
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				If Not IsNumeric(Trim(txtImageWSize(0).Text)) Then
					MsgBox("숫자만 입력이 가능합니다.", MsgBoxStyle.Information, Me.Text)
					txtImageWSize(0).Focus()
					Exit Sub
				End If
				'-- 유효성 검사 [세로Size]
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				If Trim(txtImageHSize(0).Text) = "" Then
					MsgBox("세로Size를 입력하세요.", MsgBoxStyle.Information, Me.Text)
					txtImageHSize(0).Focus()
					Exit Sub
				End If
				'-- 유효성 검사 [세로Size]
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				If Not IsNumeric(Trim(txtImageHSize(0).Text)) Then
					MsgBox("숫자만 입력이 가능합니다.", MsgBoxStyle.Information, Me.Text)
					txtImageHSize(0).Focus()
					Exit Sub
				End If
				
				'-- 동일명칭 체크
				If findSameCtrlNm(CStr(sstType.SelectedIndex), (txtTitle.Text)) Then
					MsgBox("동일한 항목명은 사용할 수 없습니다.", MsgBoxStyle.Information, Me.Text)
					Exit Sub
				End If
				
				'-- Static Image 개체만들기
				gblCtrlIdx = gblCtrlIdx + 1
				gblCtrlNm = "Control_" & gblCtrlIdx
				
				ClsEventObject = New ClassEventObject
				'Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectSImage, txtTitle.Text)
				obj = ClsEventObject.CreateObject_Renamed(Me, ClsEventMonitor, ClassEventMonitor.EventObjectID.EventObjectSImage, gblCtrlNm)
				If Not obj Is Nothing Then
					'UPGRADE_WARNING: Dir에 새 동작이 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					If Dir(txtImageName(0).Text) = "" Then
						'UPGRADE_WARNING: obj.Picture 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Picture = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\image\noimage.bmp")
					Else
						'UPGRADE_WARNING: obj.Picture 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Picture = System.Drawing.Image.FromFile(txtImageName(0).Text)
					End If
					'UPGRADE_WARNING: obj.Tag 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Tag = txtTitle.Text
					'UPGRADE_WARNING: obj.DataMember 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.DataMember = txtImageName(0).Text '-- 이미지 경로
					'UPGRADE_WARNING: obj.Stretch 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Stretch = True
					'UPGRADE_WARNING: obj.Width 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Width = CDbl(txtImageWSize(0).Text) * CDbl(gDevide)
					'UPGRADE_WARNING: obj.Height 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Height = CDbl(txtImageHSize(0).Text) * CDbl(gDevide)
					'UPGRADE_WARNING: obj.Top 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Top = CDbl(txtYpos.Text) * CDbl(gDevide)
					'UPGRADE_WARNING: obj.Left 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Left = CDbl(txtXpos.Text) * CDbl(gDevide)
					'UPGRADE_WARNING: obj.MousePointer 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.MousePointer = 5
					'UPGRADE_WARNING: obj.ToolTipText 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.ToolTipText = CStr(chkIStatic.CheckState) '-- 무조건고정
					'UPGRADE_WARNING: obj.DataField 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.DataField = IIf(chkPrint.CheckState = CDbl("1"), "0", "1") '-- 출력안함
					Call MakeSpdSaveList(obj, (sstType.SelectedIndex))
				Else
					If gblCtrlIdx > 0 Then gblCtrlIdx = gblCtrlIdx - 1
					MsgBox("동일한 항목명은 사용할 수 없습니다.", MsgBoxStyle.Information, Me.Text)
					'UPGRADE_NOTE: ClsEventObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					ClsEventObject = Nothing
					Exit Sub
				End If
				
			Case 3 '## Dynamic Image ##
				'-- 유효성 검사 [이미지명]
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				If Trim(txtImageName(1).Text) = "" Then
					MsgBox("이미지를 선택하세요.", MsgBoxStyle.Information, Me.Text)
					Call cmdImage_Click(cmdImage.Item(1), New System.EventArgs())
					Exit Sub
				End If
				'-- 유효성 검사 [가로Size]
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				If Trim(txtImageWSize(1).Text) = "" Then
					MsgBox("가로Size를 입력하세요.", MsgBoxStyle.Information, Me.Text)
					txtImageWSize(1).Focus()
					Exit Sub
				End If
				'-- 유효성 검사 [가로Size]
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				If Not IsNumeric(Trim(txtImageWSize(1).Text)) Then
					MsgBox("숫자만 입력이 가능합니다.", MsgBoxStyle.Information, Me.Text)
					txtImageWSize(1).Focus()
					Exit Sub
				End If
				'-- 유효성 검사 [세로Size]
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				If Trim(txtImageHSize(1).Text) = "" Then
					MsgBox("세로Size를 입력하세요.", MsgBoxStyle.Information, Me.Text)
					txtImageHSize(1).Focus()
					Exit Sub
				End If
				'-- 유효성 검사 [세로Size]
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				If Not IsNumeric(Trim(txtImageHSize(1).Text)) Then
					MsgBox("숫자만 입력이 가능합니다.", MsgBoxStyle.Information, Me.Text)
					txtImageHSize(1).Focus()
					Exit Sub
				End If
				
				'-- 동일명칭 체크
				If findSameCtrlNm(CStr(sstType.SelectedIndex), (txtTitle.Text)) Then
					MsgBox("동일한 항목명은 사용할 수 없습니다.", MsgBoxStyle.Information, Me.Text)
					Exit Sub
				End If
				
				'-- Dynamic Image 개체만들기
				gblCtrlIdx = gblCtrlIdx + 1
				gblCtrlNm = "Control_" & gblCtrlIdx
				
				ClsEventObject = New ClassEventObject
				'Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectDImage, txtTitle.Text)
				obj = ClsEventObject.CreateObject_Renamed(Me, ClsEventMonitor, ClassEventMonitor.EventObjectID.EventObjectDImage, gblCtrlNm)
				If Not obj Is Nothing Then
					'UPGRADE_WARNING: Dir에 새 동작이 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					If Dir(txtImageName(1).Text) = "" Then
						'UPGRADE_WARNING: obj.Picture 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Picture = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\image\noimage.bmp")
					Else
						'UPGRADE_WARNING: obj.Picture 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Picture = System.Drawing.Image.FromFile(txtImageName(1).Text)
					End If
					'UPGRADE_WARNING: obj.Tag 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Tag = txtTitle.Text
					'UPGRADE_WARNING: obj.DataMember 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.DataMember = txtImageName(1).Text '-- 이미지 경로
					'UPGRADE_WARNING: obj.Stretch 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Stretch = True
					'UPGRADE_WARNING: obj.Width 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Width = CDbl(txtImageWSize(1).Text) * CDbl(gDevide)
					'UPGRADE_WARNING: obj.Height 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Height = CDbl(txtImageHSize(1).Text) * CDbl(gDevide)
					'UPGRADE_WARNING: obj.Top 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Top = CDbl(txtYpos.Text) * CDbl(gDevide)
					'UPGRADE_WARNING: obj.Left 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Left = CDbl(txtXpos.Text) * CDbl(gDevide)
					'UPGRADE_WARNING: obj.DataField 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.DataField = IIf(chkPrint.CheckState = CDbl("1"), "0", "1") '-- 출력안함
					'UPGRADE_WARNING: obj.MousePointer 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.MousePointer = 5
					Call MakeSpdSaveList(obj, (sstType.SelectedIndex))
				Else
					If gblCtrlIdx > 0 Then gblCtrlIdx = gblCtrlIdx - 1
					MsgBox("동일한 항목명은 사용할 수 없습니다.", MsgBoxStyle.Information, Me.Text)
					'UPGRADE_NOTE: ClsEventObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					ClsEventObject = Nothing
					Exit Sub
				End If
				
			Case 4 '## Barcode ##
				'-- 유효성 검사 [길이Size]
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				If Trim(txtBarWSize.Text) = "" Then
					MsgBox("길이Size를 입력하세요.", MsgBoxStyle.Information, Me.Text)
					txtBarWSize.Focus()
					Exit Sub
				End If
				'-- 유효성 검사 [길이Size]
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				If Not IsNumeric(Trim(txtBarWSize.Text)) Then
					MsgBox("숫자만 입력이 가능합니다.", MsgBoxStyle.Information, Me.Text)
					txtBarWSize.Focus()
					Exit Sub
				End If
				'-- 유효성 검사 [높이Size]
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				If Trim(txtBarHSize.Text) = "" Then
					MsgBox("높이Size를 입력하세요.", MsgBoxStyle.Information, Me.Text)
					txtBarHSize.Focus()
					Exit Sub
				End If
				'-- 유효성 검사 [높이Size]
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				If Not IsNumeric(Trim(txtBarHSize.Text)) Then
					MsgBox("숫자만 입력이 가능합니다.", MsgBoxStyle.Information, Me.Text)
					txtBarHSize.Focus()
					Exit Sub
				End If
				'-- 유효성 검사 [높이Size]
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				If Trim(txtBarData.Text) = "" Then
					MsgBox("Data를 입력하세요.", MsgBoxStyle.Information, Me.Text)
					txtBarData.Focus()
					Exit Sub
				End If
				
				'-- 동일명칭 체크
				If findSameCtrlNm(CStr(sstType.SelectedIndex), (txtTitle.Text)) Then
					MsgBox("동일한 항목명은 사용할 수 없습니다.", MsgBoxStyle.Information, Me.Text)
					Exit Sub
				End If
				
				'-- Barcode 개체만들기
				gblCtrlIdx = gblCtrlIdx + 1
				gblCtrlNm = "Control_" & gblCtrlIdx
				
				ClsEventObject = New ClassEventObject
				'Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectBarcode, txtTitle.Text)
				obj = ClsEventObject.CreateObject_Renamed(Me, ClsEventMonitor, ClassEventMonitor.EventObjectID.EventObjectBarcode, gblCtrlNm)
				If Not obj Is Nothing Then
					'UPGRADE_WARNING: obj.Tag 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Tag = txtTitle.Text
					'UPGRADE_WARNING: obj.Caption 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Caption = txtBarData.Text
					'UPGRADE_WARNING: obj.Style 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Style = cboBarType.SelectedIndex
					'UPGRADE_WARNING: obj.Alignment 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Alignment = BarcodLib.AlignmentConstants.bcALeft
					'UPGRADE_WARNING: obj.BarWidth 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.BarWidth = 0
					'UPGRADE_WARNING: obj.Width 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Width = CDbl(txtBarWSize.Text) * CDbl(gDevide)
					'UPGRADE_WARNING: obj.Height 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Height = CDbl(txtBarHSize.Text) * CDbl(gDevide)
					'UPGRADE_WARNING: obj.Top 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Top = CDbl(txtYpos.Text) * CDbl(gDevide)
					'UPGRADE_WARNING: obj.Left 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Left = CDbl(txtXpos.Text) * CDbl(gDevide)
					'UPGRADE_WARNING: obj.Direction 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Direction = IIf(chkBarRotate.CheckState = CDbl("0"), 0, 2)
					'obj.DataField = IIf(chkPrint.Value = "1", "0", "1")         '-- 출력안함
					'UPGRADE_WARNING: obj.Visible 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Visible = False
					
					'UPGRADE_WARNING: obj.Container 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Container = Picture1
					m_ColCommandButton.Add(ClsEventObject)
					'UPGRADE_NOTE: ClsEventObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					ClsEventObject = Nothing
					
					'                If strBarImgName = "" Then
					'                    strBarImgName = txtTitle.Text & "_IMG1"
					'                Else
					'                    strBarImgName = Mid(strBarImgName, 1, Len(strBarImgName) - 1) & Right(strBarImgName, 1) + 1
					'                End If
					
					'-- 동일명칭 체크
					If findSameCtrlNm(CStr(sstType.SelectedIndex), (txtTitle.Text)) Then
						MsgBox("동일한 항목명은 사용할 수 없습니다.", MsgBoxStyle.Information, Me.Text)
						Exit Sub
					End If
					
					gblCtrlNm = gblCtrlNm & "_IMG"
					Call MakeSpdSaveList(obj, (sstType.SelectedIndex))
					
					'== 바코드를 이미지 형태로 올리기 ===================================================================
					'gblCtrlNm = gblCtrlNm & "_IMG"
					
					ClsEventObject = New ClassEventObject
					'Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectBImage, strBarImgName)
					obj = ClsEventObject.CreateObject_Renamed(Me, ClsEventMonitor, ClassEventMonitor.EventObjectID.EventObjectBImage, gblCtrlNm)
					If Not obj Is Nothing Then
						'UPGRADE_WARNING: obj.Tag 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Tag = txtTitle.Text
						If chkBarRotate.CheckState = CDbl("0") Then
							'UPGRADE_WARNING: obj.Picture 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							obj.Picture = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\" & gImage & "\barcode.bmp")
							'UPGRADE_WARNING: obj.DataMember 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							obj.DataMember = My.Application.Info.DirectoryPath & "\" & gImage & "\barcode.bmp"
							'UPGRADE_WARNING: obj.Width 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							obj.Width = CDbl(txtBarWSize.Text) * CDbl(gDevide)
							'UPGRADE_WARNING: obj.Height 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							obj.Height = CDbl(txtBarHSize.Text) * CDbl(gDevide)
						Else
							'UPGRADE_WARNING: obj.Picture 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							obj.Picture = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\" & gImage & "\barcode90.bmp")
							'UPGRADE_WARNING: obj.DataMember 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							obj.DataMember = My.Application.Info.DirectoryPath & "\" & gImage & "\barcode90.bmp"
							'UPGRADE_WARNING: obj.Width 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							obj.Width = CDbl(txtBarHSize.Text) * CDbl(gDevide)
							'UPGRADE_WARNING: obj.Height 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							obj.Height = CDbl(txtBarWSize.Text) * CDbl(gDevide)
						End If
						'UPGRADE_WARNING: obj.Stretch 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Stretch = True
						'UPGRADE_WARNING: obj.Top 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Top = CDbl(txtYpos.Text) * CDbl(gDevide)
						'UPGRADE_WARNING: obj.Left 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Left = CDbl(txtXpos.Text) * CDbl(gDevide)
						'UPGRADE_WARNING: obj.ToolTipText 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.ToolTipText = CStr(cboBarType.SelectedIndex)
						'UPGRADE_WARNING: obj.DataField 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.DataField = IIf(chkPrint.CheckState = CDbl("1"), "0", "1") '-- 출력안함
						'UPGRADE_WARNING: obj.MousePointer 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.MousePointer = 5
					Else
						If gblCtrlIdx > 0 Then gblCtrlIdx = gblCtrlIdx - 1
						MsgBox("동일한 항목명은 사용할 수 없습니다.[바코드 생성 오류]", MsgBoxStyle.Information, Me.Text)
						'UPGRADE_NOTE: ClsEventObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						ClsEventObject = Nothing
						Exit Sub
					End If
					'== 바코드를 이미지 형태로 올리기 ===================================================================
				Else
					If gblCtrlIdx > 0 Then gblCtrlIdx = gblCtrlIdx - 1
					MsgBox("동일한 항목명은 사용할 수 없습니다.", MsgBoxStyle.Information, Me.Text)
					'UPGRADE_NOTE: ClsEventObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					ClsEventObject = Nothing
					Exit Sub
				End If
				
			Case 5 '## Line ##
				'-- 유효성 검사 [선굵기]
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				If Trim(txtLineHSize.Text) = "" Then
					MsgBox("선굵기를 입력하세요.", MsgBoxStyle.Information, Me.Text)
					txtLineHSize.Focus()
					Exit Sub
				End If
				'-- 유효성 검사 [선굵기]
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				If Not IsNumeric(Trim(txtLineHSize.Text)) Then
					MsgBox("숫자만 입력이 가능합니다.", MsgBoxStyle.Information, Me.Text)
					txtLineHSize.Focus()
					Exit Sub
				End If
				'-- 유효성 검사 [선길이]
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				If Trim(txtLineWSize.Text) = "" Then
					MsgBox("선길이를 입력하세요.", MsgBoxStyle.Information, Me.Text)
					txtLineWSize.Focus()
					Exit Sub
				End If
				'-- 유효성 검사 [선길이]
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				If Not IsNumeric(Trim(txtLineWSize.Text)) Then
					MsgBox("숫자만 입력이 가능합니다.", MsgBoxStyle.Information, Me.Text)
					txtLineWSize.Focus()
					Exit Sub
				End If
				
				'-- 동일명칭 체크
				If findSameCtrlNm(CStr(sstType.SelectedIndex), (txtTitle.Text)) Then
					MsgBox("동일한 항목명은 사용할 수 없습니다.", MsgBoxStyle.Information, Me.Text)
					Exit Sub
				End If
				
				'-- Line 개체만들기
				gblCtrlIdx = gblCtrlIdx + 1
				gblCtrlNm = "Control_" & gblCtrlIdx
				
				ClsEventObject = New ClassEventObject
				obj = ClsEventObject.CreateObject_Renamed(Me, ClsEventMonitor, ClassEventMonitor.EventObjectID.EventObjectLImage, gblCtrlNm)
				If Not obj Is Nothing Then
					'UPGRADE_WARNING: obj.Tag 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.Tag = txtTitle.Text
					If chkLineRotate.CheckState = 0 Then
						'UPGRADE_WARNING: obj.Picture 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Picture = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\" & gImage & "wline.jpg")
						'UPGRADE_WARNING: obj.Stretch 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Stretch = True
						'UPGRADE_WARNING: obj.Width 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Width = CDbl(txtLineWSize.Text) * CDbl(gScaleCal)
						'UPGRADE_WARNING: obj.Height 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Height = CDbl(txtLineHSize.Text) * CDbl(gScaleCal)
						'UPGRADE_WARNING: obj.Top 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Top = CDbl(txtYpos.Text) * CDbl(gScaleCal)
						'UPGRADE_WARNING: obj.Left 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Left = CDbl(txtXpos.Text) * CDbl(gScaleCal)
						'UPGRADE_WARNING: obj.DataMember 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.DataMember = "0"
					Else
						'UPGRADE_WARNING: obj.Picture 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Picture = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\" & gImage & "hline.jpg")
						'UPGRADE_WARNING: obj.Stretch 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Stretch = True
						'UPGRADE_WARNING: obj.Width 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Width = CDbl(txtLineHSize.Text) * CDbl(gScaleCal)
						'UPGRADE_WARNING: obj.Height 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Height = CDbl(txtLineWSize.Text) * CDbl(gScaleCal)
						'UPGRADE_WARNING: obj.Top 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Top = CDbl(txtYpos.Text) * CDbl(gScaleCal)
						'UPGRADE_WARNING: obj.Left 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.Left = CDbl(txtXpos.Text) * CDbl(gScaleCal)
						'UPGRADE_WARNING: obj.DataMember 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						obj.DataMember = "1"
					End If
					'UPGRADE_WARNING: obj.MousePointer 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					obj.MousePointer = 5
					Call MakeSpdSaveList(obj, (sstType.SelectedIndex))
				Else
					If gblCtrlIdx > 0 Then gblCtrlIdx = gblCtrlIdx - 1
					MsgBox("동일한 항목명은 사용할 수 없습니다.", MsgBoxStyle.Information, Me.Text)
					'UPGRADE_NOTE: ClsEventObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					ClsEventObject = Nothing
					Exit Sub
				End If
		End Select
		
		
		'    Dim lnghNewFont As Long
		'    Dim lnghOriginalFonrt As Long
		'    Dim lngHeight As Long
		'    Dim lngWidth As Long
		'    Dim intAngle As Integer
		
		
		'UPGRADE_WARNING: obj.Visible 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj.Visible = True
		'UPGRADE_WARNING: obj.Container 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj.Container = Picture1
		
		m_ColCommandButton.Add(ClsEventObject)
		
		'UPGRADE_NOTE: ClsEventObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		ClsEventObject = Nothing
		
		'    intAngle = 90
		'    With Picture1
		'        .ScaleMode = vbPixels
		'        .AutoRedraw = True
		'        lngHeight = .TextHeight(obj)
		'        lngWidth = 0
		'
		'        With .Font
		'            lnghNewFont = CreateFont(lngHeight, lngWidth, intAngle * 10, intAngle * 10, .Weight, .Italic, .Underline, .Strikethrough, .Charset, 0, 0, 0, 0, .Name)
		'        End With
		'        lnghOriginalFonrt = SelectObject(.hdc, lnghNewFont)
		'        .CurrentX = obj.Left
		'        .CurrentY = obj.Top
		'        Picture1.Print obj
		'
		'        lnghNewFont = SelectObject(.hdc, lnghOriginalFonrt)
		'        .AutoRedraw = False
		'    End With
		'    DeleteObject lnghNewFont
		'    'obj.Visible = False
		
		
	End Sub
	
	Private Sub objSet()
		Dim strNm As String
		
		Select Case sstType.SelectedIndex
			Case 0 'Static Label
				CType(Me.Controls(txtTag.Text), Object).Tag = txtTitle.Text
				'UPGRADE_WARNING: Me.Controls().AutoSize 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CType(Me.Controls(txtTag.Text), Object).AutoSize = True
				CType(Me.Controls(txtTag.Text), Object).BackColor = System.Drawing.Color.White
				CType(Me.Controls(txtTag.Text), Object).Font = VB6.FontChangeName(CType(Me.Controls(txtTag.Text), Object).Font, txtFontName(sstType.SelectedIndex).Text)
				CType(Me.Controls(txtTag.Text), Object).Font = VB6.FontChangeSize(CType(Me.Controls(txtTag.Text), Object).Font, CDbl(txtFontSize(sstType.SelectedIndex).Text) * CDbl(gDevide))
				CType(Me.Controls(txtTag.Text), Object).Font = VB6.FontChangeBold(CType(Me.Controls(txtTag.Text), Object).Font, IIf(chkFontBold(sstType.SelectedIndex).CheckState = 1, True, False))
				CType(Me.Controls(txtTag.Text), Object).Font = VB6.FontChangeItalic(CType(Me.Controls(txtTag.Text), Object).Font, IIf(chkFontItalic(sstType.SelectedIndex).CheckState = 1, True, False))
				CType(Me.Controls(txtTag.Text), Object).Font = VB6.FontChangeUnderline(CType(Me.Controls(txtTag.Text), Object).Font, IIf(chkFontUnder(sstType.SelectedIndex).CheckState = 1, True, False))
				CType(Me.Controls(txtTag.Text), Object).Top = VB6.TwipsToPixelsY(CDbl(txtYpos.Text) * CDbl(gDevide))
				CType(Me.Controls(txtTag.Text), Object).Left = VB6.TwipsToPixelsX(CDbl(txtXpos.Text) * CDbl(gDevide))
				CType(Me.Controls(txtTag.Text), Object).Text = txtContent(sstType.SelectedIndex).Text
				'FIXIT: txtTag.Text).DataMember property 은(는) Visual Basic .NET에서 해당되는 항목이 없으므로 업그레이드되지 않습니다.     FixIT90210ae-R7593-R67265
				'UPGRADE_ISSUE: Control 메서드 Controls.DataMember이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				CType(Me.Controls(txtTag.Text), Object).DataMember = chkTStatic.CheckState
				'FIXIT: txtTag.Text).DataField property 은(는) Visual Basic .NET에서 해당되는 항목이 없으므로 업그레이드되지 않습니다.     FixIT90210ae-R7593-R67265
				'UPGRADE_ISSUE: Control 메서드 Controls.DataField이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				CType(Me.Controls(txtTag.Text), Object).DataField = IIf(chkPrint.CheckState = CDbl("1"), "0", "1") '-- 출력안함
				
			Case 1 'Dynamic Label
				CType(Me.Controls(txtTag.Text), Object).Tag = txtTitle.Text
				'UPGRADE_WARNING: Me.Controls().AutoSize 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CType(Me.Controls(txtTag.Text), Object).AutoSize = True
				CType(Me.Controls(txtTag.Text), Object).BackColor = System.Drawing.Color.White
				CType(Me.Controls(txtTag.Text), Object).Font = VB6.FontChangeName(CType(Me.Controls(txtTag.Text), Object).Font, txtFontName(sstType.SelectedIndex).Text)
				CType(Me.Controls(txtTag.Text), Object).Font = VB6.FontChangeSize(CType(Me.Controls(txtTag.Text), Object).Font, CDbl(txtFontSize(sstType.SelectedIndex).Text) * CDbl(gDevide))
				CType(Me.Controls(txtTag.Text), Object).Font = VB6.FontChangeBold(CType(Me.Controls(txtTag.Text), Object).Font, IIf(chkFontBold(sstType.SelectedIndex).CheckState = 1, True, False))
				CType(Me.Controls(txtTag.Text), Object).Font = VB6.FontChangeItalic(CType(Me.Controls(txtTag.Text), Object).Font, IIf(chkFontItalic(sstType.SelectedIndex).CheckState = 1, True, False))
				CType(Me.Controls(txtTag.Text), Object).Font = VB6.FontChangeUnderline(CType(Me.Controls(txtTag.Text), Object).Font, IIf(chkFontUnder(sstType.SelectedIndex).CheckState = 1, True, False))
				CType(Me.Controls(txtTag.Text), Object).Top = VB6.TwipsToPixelsY(CDbl(txtYpos.Text) * CDbl(gDevide))
				CType(Me.Controls(txtTag.Text), Object).Left = VB6.TwipsToPixelsX(CDbl(txtXpos.Text) * CDbl(gDevide))
				CType(Me.Controls(txtTag.Text), Object).Text = txtContent(sstType.SelectedIndex).Text
				'FIXIT: txtTag.Text).DataMember property 은(는) Visual Basic .NET에서 해당되는 항목이 없으므로 업그레이드되지 않습니다.     FixIT90210ae-R7593-R67265
				'UPGRADE_ISSUE: Control 메서드 Controls.DataMember이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				CType(Me.Controls(txtTag.Text), Object).DataMember = IIf(chkPrint.CheckState = CDbl("1"), "0", "1") '-- 출력안함
				
			Case 2 'Static Image
				CType(Me.Controls(txtTag.Text), Object).Tag = txtTitle.Text
				CType(Me.Controls(txtTag.Text), Object).Width = VB6.TwipsToPixelsX(CDbl(txtImageWSize(0).Text) * CDbl(gDevide))
				CType(Me.Controls(txtTag.Text), Object).Height = VB6.TwipsToPixelsY(CDbl(txtImageHSize(0).Text) * CDbl(gDevide))
				CType(Me.Controls(txtTag.Text), Object).Top = VB6.TwipsToPixelsY(CDbl(txtYpos.Text) * CDbl(gDevide))
				CType(Me.Controls(txtTag.Text), Object).Left = VB6.TwipsToPixelsX(CDbl(txtXpos.Text) * CDbl(gDevide))
				'UPGRADE_WARNING: Dir에 새 동작이 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				If Dir(txtImageName(0).Text) = "" Then
					'UPGRADE_ISSUE: Control 메서드 Controls.Picture이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					CType(Me.Controls(txtTag.Text), Object).Image = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\" & gImage & "noimage.bmp")
				Else
					'UPGRADE_ISSUE: Control 메서드 Controls.Picture이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					CType(Me.Controls(txtTag.Text), Object).Image = System.Drawing.Image.FromFile(txtImageName(0).Text)
				End If
				
				'FIXIT: txtTag.Text).DataMember property 은(는) Visual Basic .NET에서 해당되는 항목이 없으므로 업그레이드되지 않습니다.     FixIT90210ae-R7593-R67265
				'UPGRADE_ISSUE: Control 메서드 Controls.DataMember이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				CType(Me.Controls(txtTag.Text), Object).DataMember = txtImageName(0).Text '-- 이미지경로
				
				'FIXIT: txtTag.Text).DataField property 은(는) Visual Basic .NET에서 해당되는 항목이 없으므로 업그레이드되지 않습니다.     FixIT90210ae-R7593-R67265
				'UPGRADE_ISSUE: Control 메서드 Controls.DataField이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				CType(Me.Controls(txtTag.Text), Object).DataField = IIf(chkPrint.CheckState = CDbl("1"), "0", "1") '-- 출력안함
				
			Case 3 'Dynamic Image
				CType(Me.Controls(txtTag.Text), Object).Tag = txtTitle.Text
				CType(Me.Controls(txtTag.Text), Object).Width = VB6.TwipsToPixelsX(CDbl(txtImageWSize(1).Text) * CDbl(gDevide))
				CType(Me.Controls(txtTag.Text), Object).Height = VB6.TwipsToPixelsY(CDbl(txtImageHSize(1).Text) * CDbl(gDevide))
				CType(Me.Controls(txtTag.Text), Object).Top = VB6.TwipsToPixelsY(CDbl(txtYpos.Text) * CDbl(gDevide))
				CType(Me.Controls(txtTag.Text), Object).Left = VB6.TwipsToPixelsX(CDbl(txtXpos.Text) * CDbl(gDevide))
				
				'UPGRADE_WARNING: Dir에 새 동작이 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				If Dir(txtImageName(1).Text) = "" Then
					'UPGRADE_ISSUE: Control 메서드 Controls.Picture이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					CType(Me.Controls(txtTag.Text), Object).Image = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\" & gImage & "noimage.bmp")
				Else
					'UPGRADE_ISSUE: Control 메서드 Controls.Picture이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					CType(Me.Controls(txtTag.Text), Object).Image = System.Drawing.Image.FromFile(txtImageName(1).Text)
				End If
				
				'FIXIT: txtTag.Text).DataMember property 은(는) Visual Basic .NET에서 해당되는 항목이 없으므로 업그레이드되지 않습니다.     FixIT90210ae-R7593-R67265
				'UPGRADE_ISSUE: Control 메서드 Controls.DataMember이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				CType(Me.Controls(txtTag.Text), Object).DataMember = txtImageName(1).Text '-- 이미지경로
				'FIXIT: txtTag.Text).DataField property 은(는) Visual Basic .NET에서 해당되는 항목이 없으므로 업그레이드되지 않습니다.     FixIT90210ae-R7593-R67265
				'UPGRADE_ISSUE: Control 메서드 Controls.DataField이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				CType(Me.Controls(txtTag.Text), Object).DataField = IIf(chkPrint.CheckState = CDbl("1"), "0", "1") '-- 출력안함
				
			Case 4 'Barcode Label
				'-- 바코드 이미지 적용
				strNm = txtTag.Text
				CType(Me.Controls(txtTag.Text), Object).Tag = txtTitle.Text
				CType(Me.Controls(strNm), Object).Top = VB6.TwipsToPixelsY(CDbl(txtYpos.Text) * CDbl(gDevide))
				CType(Me.Controls(strNm), Object).Left = VB6.TwipsToPixelsX(CDbl(txtXpos.Text) * CDbl(gDevide))
				If chkBarRotate.CheckState = CDbl("0") Then
					CType(Me.Controls(strNm), Object).Width = VB6.TwipsToPixelsX(CDbl(txtBarWSize.Text) * CDbl(gDevide))
					CType(Me.Controls(strNm), Object).Height = VB6.TwipsToPixelsY(CDbl(txtBarHSize.Text) * CDbl(gDevide))
					'UPGRADE_ISSUE: Control 메서드 Controls.Picture이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					CType(Me.Controls(strNm), Object).Image = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\" & gImage & "barcode.bmp")
				Else
					CType(Me.Controls(strNm), Object).Height = VB6.TwipsToPixelsY(CDbl(txtBarWSize.Text) * CDbl(gDevide))
					CType(Me.Controls(strNm), Object).Width = VB6.TwipsToPixelsX(CDbl(txtBarHSize.Text) * CDbl(gDevide))
					'UPGRADE_ISSUE: Control 메서드 Controls.Picture이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					CType(Me.Controls(strNm), Object).Image = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\" & gImage & "barcode90.bmp")
				End If
				Me.ToolTip1.SetToolTip(CType(Me.Controls(strNm), Object), CStr(cboBarType.SelectedIndex)) '-- 바코드 타입
				'UPGRADE_ISSUE: Control 메서드 Controls.DataField이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				CType(Me.Controls(strNm), Object).DataField = IIf(chkPrint.CheckState = CDbl("1"), "0", "1") '-- 출력안함
				
				'-- 바코드 적용
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
				'FIXIT: 'Mid' 함수를 'Mid$' 함수로 바꾸십시오.                                                        FixIT90210ae-R9757-R1B8ZE
				strNm = Mid(Trim(txtTag.Text), 1, InStr(Trim(txtTag.Text), "_IMG") - 1)
				CType(Me.Controls(strNm), Object).Tag = txtTitle.Text
				CType(Me.Controls(strNm), Object).Text = txtBarData.Text
				'UPGRADE_WARNING: Me.Controls().Style 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CType(Me.Controls(strNm), Object).Style = cboBarType.SelectedIndex
				'UPGRADE_WARNING: Me.Controls().Alignment 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CType(Me.Controls(strNm), Object).Alignment = BarcodLib.AlignmentConstants.bcALeft
				CType(Me.Controls(strNm), Object).Top = VB6.TwipsToPixelsY(CDbl(txtYpos.Text) * CDbl(gDevide))
				CType(Me.Controls(strNm), Object).Left = VB6.TwipsToPixelsX(CDbl(txtXpos.Text) * CDbl(gDevide))
				If chkBarRotate.CheckState = CDbl("0") Then
					CType(Me.Controls(strNm), Object).Width = VB6.TwipsToPixelsX(CDbl(txtBarWSize.Text) * CDbl(gDevide))
					CType(Me.Controls(strNm), Object).Height = VB6.TwipsToPixelsY(CDbl(txtBarHSize.Text) * CDbl(gDevide))
				Else
					CType(Me.Controls(strNm), Object).Width = VB6.TwipsToPixelsX(CDbl(txtBarHSize.Text) * CDbl(gDevide))
					CType(Me.Controls(strNm), Object).Height = VB6.TwipsToPixelsY(CDbl(txtBarWSize.Text) * CDbl(gDevide))
				End If
				'UPGRADE_WARNING: Me.Controls().Direction 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CType(Me.Controls(strNm), Object).Direction = IIf(chkBarRotate.CheckState = CDbl("0"), 0, 2)
				
				
			Case 5 'Line Image
				CType(Me.Controls(txtTag.Text), Object).Tag = txtTitle.Text
				If chkLineRotate.CheckState = 0 Then
					CType(Me.Controls(txtTag.Text), Object).Width = VB6.TwipsToPixelsX(CDbl(txtLineWSize.Text) * CDbl(gDevide))
					CType(Me.Controls(txtTag.Text), Object).Height = VB6.TwipsToPixelsY(CDbl(txtLineHSize.Text) * CDbl(gDevide))
					CType(Me.Controls(txtTag.Text), Object).Top = VB6.TwipsToPixelsY(CDbl(txtYpos.Text) * CDbl(gDevide))
					CType(Me.Controls(txtTag.Text), Object).Left = VB6.TwipsToPixelsX(CDbl(txtXpos.Text) * CDbl(gDevide))
				Else
					CType(Me.Controls(txtTag.Text), Object).Width = VB6.TwipsToPixelsX(CDbl(txtLineHSize.Text) * CDbl(gDevide))
					CType(Me.Controls(txtTag.Text), Object).Height = VB6.TwipsToPixelsY(CDbl(txtLineWSize.Text) * CDbl(gDevide))
					CType(Me.Controls(txtTag.Text), Object).Top = VB6.TwipsToPixelsY(CDbl(txtYpos.Text) * CDbl(gDevide))
					CType(Me.Controls(txtTag.Text), Object).Left = VB6.TwipsToPixelsX(CDbl(txtXpos.Text) * CDbl(gDevide))
				End If
				Me.ToolTip1.SetToolTip(CType(Me.Controls(txtTag.Text), Object), IIf(chkPrint.CheckState = CDbl("1"), "0", "1")) '-- 출력안함
				
		End Select
		
		Dim sText As String
		sText = "Living on the edge..."
		
		'    Call DrawRotatedText(picPrint.hdc, Me.Font, 900, sText, 0, Me.ScaleY(Me.TextWidth(sText), Me.ScaleMode, vbPixels))
		
		Call SetLayout((sstType.SelectedIndex))
		
	End Sub
	
	
	
	Private Sub cmdImageDevSet_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdImageDevSet.Click
		Dim Index As Short = cmdImageDevSet.GetIndex(eventSender)
		
		If txtImageWSize(Index + 2).Text = "" Or txtImageHSize(Index + 2).Text = "" Then
			Exit Sub
		End If
		
		'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
		'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
		If Trim(txtImageDevide(Index).Text) = "" And IsNumeric(txtImageDevide(Index).Text) Then
			MsgBox("이미지 배율을 확인하세요", MsgBoxStyle.OKOnly + MsgBoxStyle.Information, Me.Text)
			'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
			txtImageDevide(Index).Focus()
			Exit Sub
		End If
		
		'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
		'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
		'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
		If Trim(txtImageWSize(Index).Text) = "" And Trim(txtImageHSize(Index).Text) = "" And IsNumeric(txtImageWSize(Index).Text) And IsNumeric(txtImageHSize(Index).Text) Then
			MsgBox("이미지 사이즈를 확인하세요", MsgBoxStyle.OKOnly + MsgBoxStyle.Information, Me.Text)
			Exit Sub
		Else
			'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
			txtImageWSize(Index).Text = CStr(System.Math.Round(CDbl(txtImageWSize(Index + 2).Text) * (CDbl(txtImageDevide(Index).Text) / 100), 0))
			'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
			txtImageHSize(Index).Text = CStr(System.Math.Round(CDbl(txtImageHSize(Index + 2).Text) * (CDbl(txtImageDevide(Index).Text) / 100), 0))
		End If
		
	End Sub
	
	' 동적 컨트롤 생성
	Private Sub cmdMake_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdMake.Click
		
		'-- Mode Set [생성]
		intMode = 3
		
		Call objNewMake()
		
		Call PaintLine()
		
	End Sub
	
	
	'FIXIT: 'Index'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                             FixIT90210ae-R1672-R1B8ZE
	Private Sub objMove(ByRef Index As Object)
		Dim intRow As Short
		'FIXIT: 'strObjType'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                        FixIT90210ae-R1672-R1B8ZE
		Dim strObjType As Object
		'FIXIT: 'strObjName'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                        FixIT90210ae-R1672-R1B8ZE
		Dim strObjName As Object
		'FIXIT: 'strObjRotate'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                      FixIT90210ae-R1672-R1B8ZE
		Dim strObjRotate As Object
		
		With spdList
			Select Case Index
				Case 0 'left   - x1 좌표
					For intRow = 1 To .MaxRows
						.Row = intRow
						Call .GetText(2, intRow, strObjType)
						Call .GetText(29, intRow, strObjName)
						
						'-- 선택이동
						If chkChoice.CheckState = CDbl("1") Then
							'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
							'UPGRADE_WARNING: strObjName 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If strObjName = Trim(txtTag.Text) Then
								'UPGRADE_WARNING: strObjType 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If strObjType = 5 Then
									If chkDetail.CheckState = 1 Then
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										.Col = 5 : .Text = CStr(CDbl(Trim(.Text)) - 1)
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										.Col = 4 : .Text = CStr(CDbl(Trim(.Text)) - 1)
									Else
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										.Col = 5 : .Text = CStr(CDbl(Trim(.Text)) - 5)
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										.Col = 4 : .Text = CStr(CDbl(Trim(.Text)) - 5)
									End If
								Else
									If chkDetail.CheckState = 1 Then
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										.Col = 4 : .Text = CStr(CDbl(Trim(.Text)) - 1)
									Else
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										.Col = 4 : .Text = CStr(CDbl(Trim(.Text)) - 5)
									End If
								End If
								'-- 라인회전[strObjRotate]이 "1" 이면 좌/우 라인이다
								'-- XI,X2를 같이 변경해 주어야 한다.
								'Call .GetText(18, intRow, strObjRotate)
								CType(Me.Controls(strObjName), Object).Left = VB6.TwipsToPixelsX(CDbl(.Text) * CDbl(gDevide))
								
							End If
						Else
							'UPGRADE_WARNING: strObjType 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If strObjType = 5 Then
								If chkDetail.CheckState = 1 Then
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									.Col = 5 : .Text = CStr(CDbl(Trim(.Text)) - 1)
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									.Col = 4 : .Text = CStr(CDbl(Trim(.Text)) - 1)
								Else
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									.Col = 5 : .Text = CStr(CDbl(Trim(.Text)) - 5)
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									.Col = 4 : .Text = CStr(CDbl(Trim(.Text)) - 5)
								End If
							Else
								If chkDetail.CheckState = 1 Then
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									.Col = 4 : .Text = CStr(CDbl(Trim(.Text)) - 1)
								Else
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									.Col = 4 : .Text = CStr(CDbl(Trim(.Text)) - 5)
								End If
							End If
							'-- 라인회전[strObjRotate]이 "1" 이면 좌/우 라인이다
							'-- XI,X2를 같이 변경해 주어야 한다.
							'Call .GetText(18, intRow, strObjRotate)
							CType(Me.Controls(strObjName), Object).Left = VB6.TwipsToPixelsX(CDbl(.Text) * CDbl(gDevide))
						End If
					Next 
				Case 1 'right  + x1 좌표
					For intRow = 1 To .MaxRows
						.Row = intRow
						Call .GetText(2, intRow, strObjType)
						Call .GetText(29, intRow, strObjName)
						'Call .GetText(18, intRow, strObjRotate)
						
						'-- 선택이동
						If chkChoice.CheckState = CDbl("1") Then
							'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
							'UPGRADE_WARNING: strObjName 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If strObjName = Trim(txtTag.Text) Then
								'UPGRADE_WARNING: strObjType 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If strObjType = 5 Then
									If chkDetail.CheckState = 1 Then
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										.Col = 5 : .Text = CStr(CDbl(Trim(.Text)) + 1)
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										.Col = 4 : .Text = CStr(CDbl(Trim(.Text)) + 1)
									Else
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										.Col = 5 : .Text = CStr(CDbl(Trim(.Text)) + 5)
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										.Col = 4 : .Text = CStr(CDbl(Trim(.Text)) + 5)
									End If
								Else
									If chkDetail.CheckState = 1 Then
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										.Col = 4 : .Text = CStr(CDbl(Trim(.Text)) + 1)
									Else
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										.Col = 4 : .Text = CStr(CDbl(Trim(.Text)) + 5)
									End If
								End If
								'-- 라인회전[strObjRotate]이 "1" 이면 좌/우 라인이다
								'-- XI,X2를 같이 변경해 주어야 한다.
								'Call .GetText(18, intRow, strObjRotate)
								CType(Me.Controls(strObjName), Object).Left = VB6.TwipsToPixelsX(CDbl(.Text) * CDbl(gDevide))
								
							End If
						Else
							'UPGRADE_WARNING: strObjType 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If strObjType = 5 Then
								If chkDetail.CheckState = 1 Then
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									.Col = 5 : .Text = CStr(CDbl(Trim(.Text)) + 1)
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									.Col = 4 : .Text = CStr(CDbl(Trim(.Text)) + 1)
								Else
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									.Col = 5 : .Text = CStr(CDbl(Trim(.Text)) + 5)
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									.Col = 4 : .Text = CStr(CDbl(Trim(.Text)) + 5)
								End If
							Else
								If chkDetail.CheckState = 1 Then
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									.Col = 4 : .Text = CStr(CDbl(Trim(.Text)) + 1)
								Else
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									.Col = 4 : .Text = CStr(CDbl(Trim(.Text)) + 5)
								End If
							End If
							'-- 라인회전[strObjRotate]이 "1" 이면 좌/우 라인이다
							'-- XI,X2를 같이 변경해 주어야 한다.
							Call .GetText(18, intRow, strObjRotate)
							CType(Me.Controls(strObjName), Object).Left = VB6.TwipsToPixelsX(CDbl(.Text) * CDbl(gDevide))
						End If
					Next 
				Case 2 'top    - y1 좌표
					For intRow = 1 To .MaxRows
						.Row = intRow
						Call .GetText(2, intRow, strObjType)
						Call .GetText(29, intRow, strObjName)
						
						'-- 선택이동
						If chkChoice.CheckState = CDbl("1") Then
							'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
							'UPGRADE_WARNING: strObjName 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If strObjName = Trim(txtTag.Text) Then
								'UPGRADE_WARNING: strObjType 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If strObjType = 5 Then
									If chkDetail.CheckState = 1 Then
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										.Col = 7 : .Text = CStr(CDbl(Trim(.Text)) - 1)
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										.Col = 6 : .Text = CStr(CDbl(Trim(.Text)) - 1)
									Else
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										.Col = 7 : .Text = CStr(CDbl(Trim(.Text)) - 5)
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										.Col = 6 : .Text = CStr(CDbl(Trim(.Text)) - 5)
									End If
								Else
									If chkDetail.CheckState = 1 Then
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										.Col = 6 : .Text = CStr(CDbl(Trim(.Text)) - 1)
									Else
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										.Col = 6 : .Text = CStr(CDbl(Trim(.Text)) - 5)
									End If
								End If
								CType(Me.Controls(strObjName), Object).Top = VB6.TwipsToPixelsY(CDbl(.Text) * CDbl(gDevide))
							End If
						Else
							'UPGRADE_WARNING: strObjType 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If strObjType = 5 Then
								If chkDetail.CheckState = 1 Then
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									.Col = 7 : .Text = CStr(CDbl(Trim(.Text)) - 1)
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									.Col = 6 : .Text = CStr(CDbl(Trim(.Text)) - 1)
								Else
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									.Col = 7 : .Text = CStr(CDbl(Trim(.Text)) - 5)
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									.Col = 6 : .Text = CStr(CDbl(Trim(.Text)) - 5)
								End If
							Else
								If chkDetail.CheckState = 1 Then
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									.Col = 6 : .Text = CStr(CDbl(Trim(.Text)) - 1)
								Else
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									.Col = 6 : .Text = CStr(CDbl(Trim(.Text)) - 5)
								End If
							End If
							CType(Me.Controls(strObjName), Object).Top = VB6.TwipsToPixelsY(CDbl(.Text) * CDbl(gDevide))
						End If
					Next 
				Case 3 'bottom + y1 좌표
					For intRow = 1 To .MaxRows
						.Row = intRow
						Call .GetText(2, intRow, strObjType)
						Call .GetText(29, intRow, strObjName)
						
						'-- 선택이동
						If chkChoice.CheckState = CDbl("1") Then
							'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
							'UPGRADE_WARNING: strObjName 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If strObjName = Trim(txtTag.Text) Then
								'UPGRADE_WARNING: strObjType 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If strObjType = 5 Then
									If chkDetail.CheckState = 1 Then
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										.Col = 7 : .Text = CStr(CDbl(Trim(.Text)) + 1)
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										.Col = 6 : .Text = CStr(CDbl(Trim(.Text)) + 1)
									Else
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										.Col = 7 : .Text = CStr(CDbl(Trim(.Text)) + 5)
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										.Col = 6 : .Text = CStr(CDbl(Trim(.Text)) + 5)
									End If
								Else
									If chkDetail.CheckState = 1 Then
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										.Col = 6 : .Text = CStr(CDbl(Trim(.Text)) + 1)
									Else
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
										.Col = 6 : .Text = CStr(CDbl(Trim(.Text)) + 5)
									End If
								End If
								CType(Me.Controls(strObjName), Object).Top = VB6.TwipsToPixelsY(CDbl(.Text) * CDbl(gDevide))
							End If
						Else
							'UPGRADE_WARNING: strObjType 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If strObjType = 5 Then
								If chkDetail.CheckState = 1 Then
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									.Col = 7 : .Text = CStr(CDbl(Trim(.Text)) + 1)
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									.Col = 6 : .Text = CStr(CDbl(Trim(.Text)) + 1)
								Else
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									.Col = 7 : .Text = CStr(CDbl(Trim(.Text)) + 5)
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									.Col = 6 : .Text = CStr(CDbl(Trim(.Text)) + 5)
								End If
							Else
								If chkDetail.CheckState = 1 Then
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									.Col = 6 : .Text = CStr(CDbl(Trim(.Text)) + 1)
								Else
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
									.Col = 6 : .Text = CStr(CDbl(Trim(.Text)) + 5)
								End If
							End If
							CType(Me.Controls(strObjName), Object).Top = VB6.TwipsToPixelsY(CDbl(.Text) * CDbl(gDevide))
						End If
					Next 
				Case 4
					'-- X1,Y1 좌표설정
					For intRow = 1 To .MaxRows
						.Row = intRow
						Call .GetText(2, intRow, strObjType)
						Call .GetText(29, intRow, strObjName)
						'
						'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
						'UPGRADE_WARNING: strObjName 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: strObjType 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If strObjType = sstType.SelectedIndex And strObjName = Trim(txtTag.Text) Then
							'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
							.Col = 4 : .Text = Trim(txtXpos.Text)
							CType(Me.Controls(strObjName), Object).Left = VB6.TwipsToPixelsX(CDbl(.Text) * CDbl(gDevide))
							'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
							.Col = 6 : .Text = Trim(txtYpos.Text)
							CType(Me.Controls(strObjName), Object).Top = VB6.TwipsToPixelsY(CDbl(.Text) * CDbl(gDevide))
							Exit For
						End If
					Next 
			End Select
		End With
		
	End Sub
	
	Private Sub cmdMove_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles cmdMove.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = cmdMove.GetIndex(eventSender)
		
		'-- Mode Set [이동]
		intMode = 2
		
		Call objMove(Index)
		
		If Index < 4 Then
			intMoveIdx = Index
			
			If chkContinue.CheckState = 1 Then
				tmrMove.Interval = 100
				tmrMove.Enabled = True
				System.Windows.Forms.Application.DoEvents()
			Else
				tmrMove.Enabled = False
			End If
		End If
		
	End Sub
	
	Private Sub cmdMove_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles cmdMove.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = cmdMove.GetIndex(eventSender)
		
		tmrMove.Enabled = False
		
	End Sub
	
    Private Sub cmdPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPrint.Click
        Dim Printer As Printing.PrintForm

        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
        Dim prtSelectPrinter As Printer
        Dim boolPrinter_Select_Fales As Boolean
        Dim Buffer As String
        Dim aryPrinter() As String
        Dim strBuffer As String

        'Printer 개체를 이용해 인쇄물을 작성하실 때에는 다음의 사항을 기억하여 주십시오.
        '
        'PaperSize 는 Printer Driver에 따라 다르지만 기본적으로 A4 용지로 지정되어 있습니다.
        '용지의 크기를 사용자 정의의 크기로 지정하기 위하여 값을 256 으로 지정할 수 있지만
        '용지의 크기를 지정하는 것은 무의미합니다. 게다가 256으로 변경할 때 오류를 리터하는
        '드라이버들도 더러 있기 때문입니다.
        '용지의 크기를 지정할 필요는 없으며 인쇄물의 크기만 신경쓰시면 되겠습니다.
        '
        '님께서 적어놓은 코드를 보자면 가로 190, 세로 134 mm 의 용지에 맞게 출력을 하실려고
        '하는 것 같습니다.
        '이럴 경우 용지의 크기는 190 * 134 보다 작지만 않다면 어떤 용지규격으로 셋팅해도 관계
        '없습니다. 이럴 경우에는 그냥 A4 로 셋팅하셔도 됩니다.
        'Printer의 Width속성과 Height속성은 Twip 단위로 되어 있으며 현재 인쇄가능한 인쇄물의
        '테두리(한계, Boundary)개념으로 생각하시는 게 좋을 듯 합니다.

        '인쇄할 때 가장 중요한 것은 ScaleMode, Scale, ScaleWidth, ScaleHeight 입니다.
        '
        'mm 단위측정값을 기준으로 출력하시고자 한다면 ScaleMode 를  6 으로 지정하시면 됩니다.
        '주의할 점은 용지를 A4로, ScaleMode를 6 으로 셋팅한 후에
        'Printer.Line (0, 0)-(210, 297), , B
        '위의 문을 실행했을 경우 우측과 하단의 테두리는 범위가 초과하여 출력이 되지 않습니다.
        '왜냐하면 용지의 크기는 210 * 297 이지만 프린터마다 인쇄가능영역이라는 게 존재합니다.
        '잉크젯인 경우에는 레이저젯에 비해 같은 용지에 대해 인쇄가능영역이 작습니다.
        '그래서 ScaleMode 를 6으로 했을 때 ScaleWidth 나 ScaleHeight의 값을 보면 210 또는 297 보다
        '작은 값으로 되어 있다는 것을 알 수 있습니다.
        '이런 부분들을 고려하여 인쇄물을 작성해 보시기 바랍니다.
        '그럼 즐프~~하세요.



        ''    '============== 이미지 출력 방식 ==========================================================
        ''    Picture1.AutoRedraw = True
        ''    SendMessage Picture1.hwnd, WM_PAINT, Picture1.hDC, 0
        ''    'SendMessage Picture1.hwnd, WM_PRINT, Picture1.hDC, PRF_CHILDREN Or PRF_CLIENT Or PRF_OWNED
        ''    Printer.PaintPicture Picture1.Image, 0, 0, Picture1.Width, Picture1.Height
        ''    Printer.EndDoc
        ''    SavePicture Picture1.Image, "C:\TEST.BMP"

        ''    '============== 이미지 출력 방식 ==========================================================

        'Exit Sub

        Dim intRow As Short
        Dim intCol As Short
        Dim intCnt As Short
        'FIXIT: 'strX1' and 'strX2' and 'strY1'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                     FixIT90210ae-R1672-R1B8ZE
        Dim strX2, strX1, strY1 As Object
        Dim strY2 As String
        Dim strFont As String
        Dim strFontSize As String
        Dim strFontBold As String
        Dim strFontUnder As String
        Dim strFontItalic As String
        Dim strdata As String
        Dim strTitle As String
        Dim strPrtYN As String
        Dim intPixeltoTwip As Integer
        Dim intPixeltoTwipX As Integer
        Dim intPixeltoTwipY As Integer
        'FIXIT: 'varTmp'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                            FixIT90210ae-R1672-R1B8ZE
        Dim varTmp As Object

        If chkCorrect.CheckState = CDbl("1") Then
            '        Call spdList.GetText(23, 1, varTmp): intPixeltoTwip = IIf(varTmp <> "", varTmp, 15)
            '        Call spdList.GetText(23, 1, varTmp): intPixeltoTwipX = IIf(varTmp <> "", varTmp, 15)
            '        Call spdList.GetText(24, 1, varTmp): intPixeltoTwipX = IIf(varTmp <> "", varTmp, 15)

            intPixeltoTwip = CInt(gBojung) '14.405
            intPixeltoTwipX = CInt(gBojung) '14.405
            intPixeltoTwipY = CInt(gBojung) '14.405
        Else
            intPixeltoTwip = 15
            intPixeltoTwipX = 15
            intPixeltoTwipY = 15
        End If

        '-- 선택된 프린터로 출력
        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
        For Each prtSelectPrinter In Printers
            'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            'FIXIT: 'UCase' 함수를 'UCase$' 함수로 바꾸십시오.                                                    FixIT90210ae-R9757-R1B8ZE
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            'FIXIT: 'UCase' 함수를 'UCase$' 함수로 바꾸십시오.                                                    FixIT90210ae-R9757-R1B8ZE
            If UCase(Trim(prtSelectPrinter.DeviceName)) = UCase(Trim(cmbPrinter.Text)) Then
                'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                Printer = prtSelectPrinter
                boolPrinter_Select_Fales = True
                Exit For
            End If
        Next prtSelectPrinter

        Dim W, X, Y, H As Object
        With spdList
            'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
            Printer.ScaleMode = ScaleModeConstants.vbTwips
            'FIXIT: Picture1.AutoRedraw property 은(는) Visual Basic .NET에서 해당되는 항목이 없으므로 업그레이드되지 않습니다.     FixIT90210ae-R7593-R67265
            'UPGRADE_ISSUE: PictureBox 속성 Picture1.AutoRedraw이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            Picture1.AutoRedraw = True
            '-- 박스 그리기

            For intRow = 1 To .MaxRows
                .Row = intRow
                .Col = 2
                'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                Select Case Trim(.Text)
                    Case "0"
                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        Printer.ScaleMode = ScaleModeConstants.vbTwips
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 4
                        'UPGRADE_WARNING: strX1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        strX1 = CDbl(Trim(.Text)) * intPixeltoTwip
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 5
                        'UPGRADE_WARNING: strX2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        strX2 = CDbl(Trim(.Text)) * intPixeltoTwip
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 6
                        'UPGRADE_WARNING: strY1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        strY1 = CDbl(Trim(.Text)) * intPixeltoTwip
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 7 : strY2 = CStr(CDbl(Trim(.Text)) * intPixeltoTwip)

                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 8 : strFont = Trim(.Text)
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 9 : strFontSize = Trim(.Text)
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 10 : strFontBold = Trim(.Text)
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 11 : strFontItalic = Trim(.Text)
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 12 : strFontUnder = Trim(.Text)
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 22 : strdata = Trim(.Text)

                        'txtContentU(0).Text = strData

                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        Printer.FontName = strFont
                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        Printer.Font = VB6.FontChangeSize(Printer.Font, CDec(strFontSize))
                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        Printer.Font = VB6.FontChangeBold(Printer.Font, IIf(strFontBold = "1", True, False))
                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        Printer.Font = VB6.FontChangeItalic(Printer.Font, IIf(strFontItalic = "1", True, False))
                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        Printer.Font = VB6.FontChangeUnderline(Printer.Font, IIf(strFontUnder = "1", True, False))

                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        'UPGRADE_WARNING: strX1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Printer.CurrentX = strX1
                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        'UPGRADE_WARNING: strY1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Printer.CurrentY = strY1
                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        Printer.Print(strdata)

                        '    Picture1.Font = "Calibri"
                        '    Dim dY As Long
                        '    dY = 1
                        '    TextBox1.Text = ucs2

                        'Picture1.FontName = strFont
                        'Call TextOutW(Printer.hdc, strX1 * 15, strX2 * 15, StrPtr(txtContentU(0).Text), Len(txtContentU(0).Text))


                    Case "1"
                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        Printer.ScaleMode = ScaleModeConstants.vbTwips
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 4
                        'UPGRADE_WARNING: strX1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        strX1 = CDbl(Trim(.Text)) * intPixeltoTwip
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 5
                        'UPGRADE_WARNING: strX2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        strX2 = CDbl(Trim(.Text)) * intPixeltoTwip
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 6
                        'UPGRADE_WARNING: strY1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        strY1 = CDbl(Trim(.Text)) * intPixeltoTwip
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 7 : strY2 = CStr(CDbl(Trim(.Text)) * intPixeltoTwip)

                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 8 : strFont = Trim(.Text)
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 9 : strFontSize = Trim(.Text)
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 10 : strFontBold = Trim(.Text)
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 11 : strFontItalic = Trim(.Text)
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 12 : strFontUnder = Trim(.Text)
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 22 : strdata = Trim(.Text)

                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        Printer.FontName = strFont
                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        Printer.Font = VB6.FontChangeSize(Printer.Font, CDec(strFontSize))
                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        Printer.Font = VB6.FontChangeBold(Printer.Font, IIf(strFontBold = "1", True, False))
                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        Printer.Font = VB6.FontChangeItalic(Printer.Font, IIf(strFontItalic = "1", True, False))
                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        Printer.Font = VB6.FontChangeUnderline(Printer.Font, IIf(strFontUnder = "1", True, False))

                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        'UPGRADE_WARNING: strX1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Printer.CurrentX = strX1
                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        'UPGRADE_WARNING: strY1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Printer.CurrentY = strY1
                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        Printer.Print(strdata)

                    Case "2"
                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        Printer.ScaleMode = ScaleModeConstants.vbTwips
                        '.Col = 3: strTitle = Trim(.Text)
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 29 : strTitle = Trim(.Text)

                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 4
                        'UPGRADE_WARNING: strX1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        strX1 = CDbl(Trim(.Text)) * intPixeltoTwip
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 5
                        'UPGRADE_WARNING: strX2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        strX2 = CDbl(Trim(.Text)) * intPixeltoTwip
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 6
                        'UPGRADE_WARNING: strY1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        strY1 = CDbl(Trim(.Text)) * intPixeltoTwip
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 7 : strY2 = CStr(CDbl(Trim(.Text)) * intPixeltoTwip)

                        '                    .Col = 8: strFont = Trim(.Text)
                        '                    .Col = 9: strFontSize = Trim(.Text)
                        '                    .Col = 17: strData = Trim(.Text)

                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        Printer.PaintPicture(CType(Me.Controls(strTitle), Object), strX1, strY1, strX2, strY2)

                    Case "3"
                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        Printer.ScaleMode = ScaleModeConstants.vbTwips

                        '.Col = 3: strTitle = Trim(.Text)
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 29 : strTitle = Trim(.Text)

                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 4
                        'UPGRADE_WARNING: strX1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        strX1 = CDbl(Trim(.Text)) * intPixeltoTwip
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 5
                        'UPGRADE_WARNING: strX2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        strX2 = CDbl(Trim(.Text)) * intPixeltoTwip
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 6
                        'UPGRADE_WARNING: strY1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        strY1 = CDbl(Trim(.Text)) * intPixeltoTwip
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 7 : strY2 = CStr(CDbl(Trim(.Text)) * intPixeltoTwip)

                        '                    .Col = 8: strFont = Trim(.Text)
                        '                    .Col = 9: strFontSize = Trim(.Text)
                        '                    .Col = 17: strData = Trim(.Text)

                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        Printer.PaintPicture(CType(Me.Controls(strTitle), Object), strX1, strY1, strX2, strY2)

                    Case "4"
                        '.Col = 3: strTitle = Trim(.Text)
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 29 : strTitle = Trim(.Text)
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Mid' 함수를 'Mid$' 함수로 바꾸십시오.                                                        FixIT90210ae-R9757-R1B8ZE
                        strTitle = Mid(Trim(strTitle), 1, InStr(Trim(strTitle), "_IMG") - 1)


                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 4
                        'UPGRADE_WARNING: strX1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        strX1 = CDbl(Trim(.Text)) * intPixeltoTwip
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 5
                        'UPGRADE_WARNING: strX2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        strX2 = CDbl(Trim(.Text)) * intPixeltoTwip
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 6
                        'UPGRADE_WARNING: strY1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        strY1 = CDbl(Trim(.Text)) * intPixeltoTwip
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 7 : strY2 = CStr(CDbl(Trim(.Text)) * intPixeltoTwip)

                        'FIXIT: 'x' and 'y' and 'W' and 'H'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                         FixIT90210ae-R1672-R1B8ZE

                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        Printer.ScaleMode = ScaleModeConstants.vbTwips
                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        Printer.PSet(0, 0, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White))

                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        'UPGRADE_WARNING: strX1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: x 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        X = Printer.ScaleX(strX1, ScaleModeConstants.vbTwips) ' X-position = 25 mm from left border
                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        'UPGRADE_WARNING: strY1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: y 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Y = Printer.ScaleY(strY1, ScaleModeConstants.vbTwips) ' Y-position = 25 mm from top border
                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        'UPGRADE_WARNING: strX2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: W 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        W = Printer.ScaleX(strX2, ScaleModeConstants.vbTwips) ' Width = 100 mm
                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        'UPGRADE_WARNING: H 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        H = Printer.ScaleY(CSng(strY2), ScaleModeConstants.vbTwips) ' Height = 40 mm

                        '-- 바코드 회전
                        .Col = 16
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'UPGRADE_WARNING: Me.Controls(strTitle).Direction 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        CType(Me.Controls(strTitle), Object).Direction = IIf(Trim(.Text) = "0", 0, 2)
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        If Trim(.Text) = "0" Then
                            'UPGRADE_WARNING: Me.Controls().PrinterWidth 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: W 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            CType(Me.Controls(strTitle), Object).PrinterWidth = W '(W * 5)  'W
                            'UPGRADE_WARNING: Me.Controls().PrinterHeight 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: H 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            CType(Me.Controls(strTitle), Object).PrinterHeight = H '(H * 5)  'H
                        Else
                            'UPGRADE_WARNING: Me.Controls().PrinterWidth 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: H 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            CType(Me.Controls(strTitle), Object).PrinterWidth = H '(W * 5)  'W
                            'UPGRADE_WARNING: Me.Controls().PrinterHeight 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: W 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            CType(Me.Controls(strTitle), Object).PrinterHeight = W '(H * 5)  'H
                        End If
                        'UPGRADE_WARNING: Me.Controls().PrinterScaleMode 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_ISSUE: vbTwips 상수가 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
                        CType(Me.Controls(strTitle), Object).PrinterScaleMode = vbTwips '3:픽셀,1:트윕,6:밀리미터
                        'UPGRADE_WARNING: Me.Controls().Alignment 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        CType(Me.Controls(strTitle), Object).Alignment = BarcodLib.AlignmentConstants.bcACenter
                        'UPGRADE_WARNING: Me.Controls().PrinterLeft 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: x 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        CType(Me.Controls(strTitle), Object).PrinterLeft = X '* 4.6
                        'UPGRADE_WARNING: Me.Controls().PrinterTop 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: y 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        CType(Me.Controls(strTitle), Object).PrinterTop = Y '* 5
                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        'UPGRADE_WARNING: Me.Controls().PrinterHDC 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_ISSUE: Printer 속성 Printer.hdc이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                        CType(Me.Controls(strTitle), Object).PrinterHDC = Printer.hdc

                    Case "5"
                        '-- 출력여부
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 21 : strPrtYN = Trim(.Text)
                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        Printer.ScaleMode = ScaleModeConstants.vbTwips

                        'If strPrtYN = "1" Then

                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        Printer.PSet(0, 0, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White))

                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 4
                        'UPGRADE_WARNING: strX1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        strX1 = CDbl(Trim(.Text)) * intPixeltoTwip '* 13.3
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 5
                        'UPGRADE_WARNING: strX2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        strX2 = CDbl(Trim(.Text)) * intPixeltoTwip '* 13.3
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 6
                        'UPGRADE_WARNING: strY1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        strY1 = CDbl(Trim(.Text)) * intPixeltoTwip '* 13.3
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        .Col = 7 : strY2 = CStr(CDbl(Trim(.Text)) * intPixeltoTwip) '* 13.3
                        '선굵기
                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        Printer.DrawWidth = 1
                        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                        Printer.Line(strX1, strY1, strX2, strY2)
                        'End If
                End Select
            Next
        End With


        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
        Printer.EndDoc()

        'SavePicture Picture1.Image, "C:\TEST.BMP"

    End Sub
	
	Public Sub cmdSet_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSet.Click
		
		'-- Mode Set [적용가능]
		If intMode = 1 Then
			Call objSet()
		End If
		
	End Sub
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' 동적 버튼 생성
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'''Private Sub Command1_Click()
	'''
	'''    Dim obj                 As Object
	'''    Dim i                   As Integer
	'''    Dim ClsEventObject      As ClassEventObject
	'''
	'''    ' 프로그램 정보 TextBox 숨김
	'''    Text1.Visible = False
	'''
	'''    List1.Clear
	'''
	'''    ' 컬렉션 초기화
	''''    Set m_ColCommandButton = Nothing
	''''    Set m_ColCommandButton = New Collection
	'''
	'''    ' 동적 컨트롤 생성
	'''    For i = 1 To Val(Combo1.Text)
	'''        Set ClsEventObject = New ClassEventObject
	'''
	'''        If Option1.Value = True Then
	'''            ' CommandButton
	'''            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectCommandButton, "DynamicCmd" & CStr(i))
	'''            obj.Width = 3600
	'''            obj.Height = 525
	'''            obj.Top = 300 + (i - 1) * (525 + 30)
	'''            obj.Left = 450
	'''            obj.Caption = "Button" & CStr(i)
	'''        ElseIf Option2.Value = True Then
	'''            ' TextBox
	'''            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectTextBox, "DynamicTxt" & CStr(i))
	'''            obj.Width = 3600
	'''            obj.Height = 525
	'''            obj.Top = 300 + (i - 1) * (525 + 30)
	'''            obj.Left = 450
	'''            obj.Text = "Text" & CStr(i)
	'''        ElseIf Option3.Value = True Then
	'''            ' Label
	'''            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectLabel, "DynamicLbl" & CStr(i))
	'''            obj.Width = 3600
	'''            obj.Height = 525
	'''            obj.Top = 300 + (i - 1) * (525 + 30)
	'''            obj.Left = 450
	'''            obj.Caption = "Label" & CStr(i)
	'''        ElseIf Option4.Value = True Then
	'''            ' Image
	'''            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectImage, "DynamicImg" & CStr(i))
	'''            obj.Width = 3600
	'''            obj.Height = 525
	'''            obj.Top = 300 + (i - 1) * (525 + 30)
	'''            obj.Left = 450
	'''            obj.Picture = LoadPicture(App.Path & "\ugc.jpg")
	'''
	'''        ElseIf Option5.Value = True Then
	'''            ' line
	'''            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectLine, "DynamicLine" & CStr(i))
	'''            '-- 세로선
	'''            obj.X1 = 100 * i
	'''            obj.X2 = 100 * i
	'''            obj.Y1 = 2070
	'''            obj.Y2 = 4560
	'''            '-- 가로선
	'''            obj.X1 = 2850
	'''            obj.X2 = 7080
	'''            obj.Y1 = 100 * i
	'''            obj.Y2 = 100 * i
	'''
	'''        Else
	'''            ' barcode
	'''            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectBarcode, "DynamicBar" & CStr(i))
	'''            obj.Alignment = bcACenter
	'''            obj.Caption = "88006611"
	'''            obj.Style = msSCode128B
	'''            obj.Width = 3600
	'''            obj.Height = 525
	'''            obj.Top = 300 + (i - 1) * (525 + 30)
	'''            obj.Left = 450
	'''
	''''            Barcod1.Alignment = bcACenter
	'''            'Barcod1.Style = msSCode128B ' msS2of5
	'''
	'''        End If
	'''
	'''        obj.Visible = True
	'''        'Set obj.Container = Frame2
	'''        Set obj.Container = Picture1
	'''
	'''        m_ColCommandButton.Add ClsEventObject
	'''
	'''        Set ClsEventObject = Nothing
	'''    Next
	'''
	'''End Sub
	
	
    Private Sub MDIForm_Tool()

        On Error GoTo ErrorRouten

        With tlbMain
            'UPGRADE_ISSUE: MSComctlLib.Toolbar 속성 tlbMain.AllowCustomize이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            .AllowCustomize = False
            .ImageList = imlToolbar
            '.TextAlignment = tbrTextAlignBottom '= tbrTextAlignRight
            'UPGRADE_ISSUE: MSComctlLib.Toolbar 속성 tlbMain.TextAlignment이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            .TextAlignment = System.Windows.Forms.ToolBarTextAlign.Right
            'UPGRADE_ISSUE: MSComctlLib.Toolbar 속성 tlbMain.BorderStyle이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            .BorderStyle = System.Windows.Forms.BorderStyle.None
            'UPGRADE_ISSUE: MSComctlLib.Toolbar 속성 tlbMain.Appearance이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            .Appearance = System.Windows.Forms.BorderStyle.Fixed3D
            'UPGRADE_ISSUE: MSComctlLib.Toolbar 속성 tlbMain.Style이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            .Style = System.Windows.Forms.ToolBarAppearance.Flat
            'UPGRADE_WARNING: MSComctlLib.Buttons 메서드 tlbMain.Buttons.Add에 새 동작이 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            Call .Items.Add(New System.Windows.Forms.ToolStripButton(, TLBKEY_NEW, "신규", System.Windows.Forms.ToolBarButtonStyle.PushButton, "New"))
            'UPGRADE_WARNING: MSComctlLib.Buttons 메서드 tlbMain.Buttons.Add에 새 동작이 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            Call .Items.Add(New System.Windows.Forms.ToolStripButton(, TLBKEY_OPEN, "열기", System.Windows.Forms.ToolBarButtonStyle.PushButton, "Open"))
            'UPGRADE_WARNING: MSComctlLib.Buttons 메서드 tlbMain.Buttons.Add에 새 동작이 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            Call .Items.Add(New System.Windows.Forms.ToolStripButton(, TLBKEY_SAVE, "저장", System.Windows.Forms.ToolBarButtonStyle.PushButton, "Save"))

            'UPGRADE_WARNING: MSComctlLib.Buttons 메서드 tlbMain.Buttons.Add에 새 동작이 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            Call .Items.Add(New System.Windows.Forms.ToolStripButton(, "", "", System.Windows.Forms.ToolBarButtonStyle.Separator))

            'UPGRADE_WARNING: MSComctlLib.Buttons 메서드 tlbMain.Buttons.Add에 새 동작이 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            Call .Items.Add(New System.Windows.Forms.ToolStripButton(, TLBKEY_MAKE, "JOB", System.Windows.Forms.ToolBarButtonStyle.PushButton, "Make"))
            'UPGRADE_WARNING: MSComctlLib.Buttons 메서드 tlbMain.Buttons.Add에 새 동작이 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            Call .Items.Add(New System.Windows.Forms.ToolStripButton(, TLBKEY_VIEW, "보기", System.Windows.Forms.ToolBarButtonStyle.PushButton, "View"))
            'UPGRADE_WARNING: MSComctlLib.Buttons 메서드 tlbMain.Buttons.Add에 새 동작이 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            Call .Items.Add(New System.Windows.Forms.ToolStripButton(, "", "", System.Windows.Forms.ToolBarButtonStyle.Separator))
            'UPGRADE_WARNING: MSComctlLib.Buttons 메서드 tlbMain.Buttons.Add에 새 동작이 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            Call .Items.Add(New System.Windows.Forms.ToolStripButton(, TLBKEY_EDIT, "설정", System.Windows.Forms.ToolBarButtonStyle.PushButton, "Edit"))
            'UPGRADE_WARNING: MSComctlLib.Buttons 메서드 tlbMain.Buttons.Add에 새 동작이 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            Call .Items.Add(New System.Windows.Forms.ToolStripButton(, TLBKEY_EXIT, "종료", System.Windows.Forms.ToolBarButtonStyle.PushButton, "Exit"))
            'UPGRADE_WARNING: MSComctlLib.Buttons 메서드 tlbMain.Buttons.Add에 새 동작이 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            Call .Items.Add(New System.Windows.Forms.ToolStripButton(, "", "", System.Windows.Forms.ToolBarButtonStyle.Separator))


            .Refresh()
        End With

        Exit Sub

ErrorRouten:
        '    Call ErrMsgProc(CallForm)

    End Sub


    'Private Sub Command2_Click()
    '    Dim i As Integer
    '    Dim sTmp As String
    '    Text1.Text = "가나(다)라"
    '
    '    Picture1.Cls
    '    For i = 1 To Len(Text1.Text)
    '        If Mid(Text1.Text, i, 1) = "(" Then
    '            sTmp = Mid(Text1.Text, i, 3)
    '            i = i + 2
    '        Else
    '            sTmp = Mid(Text1.Text, i, 1)
    '        End If
    '        Picture1.CurrentX = (Picture1.ScaleWidth - Picture1.TextWidth(sTmp)) / 2
    '        Picture1.Print sTmp
    '    Next i
    '
    '
    '
    'End Sub


    Private Sub cmdUndo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdUndo.Click
        'FIXIT: 'Moveobj'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                           FixIT90210ae-R1672-R1B8ZE
        Dim Moveobj As Object
        'FIXIT: 'x'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                                 FixIT90210ae-R1672-R1B8ZE
        Dim x As Object
        Dim y As Integer

        'UPGRADE_WARNING: LMousePos.obj 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Moveobj 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Moveobj = LMousePos.obj
        'UPGRADE_WARNING: x 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        x = LMousePos.fromx
        y = LMousePos.fromy

        'UPGRADE_WARNING: x 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CType(Me.Controls(Moveobj), Object).Left = VB6.TwipsToPixelsX(x)
        CType(Me.Controls(Moveobj), Object).Top = VB6.TwipsToPixelsY(y)

    End Sub



    Private Sub Frame12_DragDrop(ByRef Source As System.Windows.Forms.Control, ByRef x As Single, ByRef y As Single)

    End Sub

    Private Sub lblPrint_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblPrint.DoubleClick

        If chkCorrect.Visible = True Then
            chkCorrect.Visible = False
        Else
            chkCorrect.Visible = True
        End If

    End Sub

    'Private Sub Command3_Click()
    '
    '    Call RotateControl(Me.Controls("Control_1"), 90)
    '
    'End Sub

    'Private Sub Form_Activate()
    '    MDIActiveX.WindowState = ccMaximize
    'End Sub
    '
    'Private Sub Form_Deactivate()
    '    MDIActiveX.WindowState = ccMinimize
    'End Sub

    Private Sub lblTitle_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblTitle.DoubleClick

        If txtTag.Visible = True Then
            txtTag.Visible = False
        Else
            txtTag.Visible = True
        End If

    End Sub

    Public Sub mnuClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuClose.Click

        If MsgBox("종료하시겠습니까?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, Me.Text) = MsgBoxResult.Yes Then
            Me.Close()
        End If

    End Sub

    Public Sub mnuMake_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuMake.Click

        If MsgBox("작업파일을 생성하시겠습니까?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, Me.Text) = MsgBoxResult.Yes Then
            Call MakeJOB()
        End If

    End Sub


    ' 첫번째 방법 : UTF-16을 나타내는 Byte Order Mark(BOM) 가 없을 경우,
    '
    Public Function UTF8FromUTF16(ByRef abytUTF16() As Byte) As Byte()

        Dim lngByteNum As Integer
        Dim abytUTF8() As Byte
        Dim lngCharCount As Integer

        On Error GoTo ConversionErr

        lngCharCount = (UBound(abytUTF16) + 1) \ 2
        ' UTF-16 LE 스트링의 문자의 수를 대입시켜, 변환에 필요한 바이트 수를 구합니다.
        lngByteNum = WideCharToMultiByteArray(CP_UTF8, 0, abytUTF16(0), lngCharCount, 0, 0, 0, 0)

        If lngByteNum > 0 Then
            ' 변환된 코드를 반환받을 메모리를 확보한 후 함수를 호출합니다.
            ReDim abytUTF8(lngByteNum - 1)
            lngByteNum = WideCharToMultiByteArray(CP_UTF8, 0, abytUTF16(0), lngCharCount, abytUTF8(0), lngByteNum, 0, 0)
            UTF8FromUTF16 = VB6.CopyArray(abytUTF8)
        End If
        Exit Function

ConversionErr:
        MsgBox(" Conversion failed ")

    End Function


    ' 두번째 방법 : BOM 을 무시한 후, UTF-8 방식으로 변환한 후,
    '                    UTF-8 방식을 나타내는 Signature 를 추가하여 반환
    '
    Public Function UTF8FromUTF16withMark(ByRef abytUTF16() As Byte) As Byte()
        Dim abytTemp() As Byte
        Dim abytUTF8() As Byte
        Dim lngByteNum As Integer
        Dim lngCharCount As Integer
        Dim lngUpper As Integer

        On Error GoTo ConversionErr

        abytTemp = VB6.CopyArray(abytUTF16)
        lngUpper = UBound(abytTemp)
        If lngUpper > 1 Then
            ' UTF-16 LE 의 바이트순서표식이 있을 경우 이를 일단 삭제합니다.
            ' &HFEFF 문자인데, LE에서는 도치되어 저장되므로, &HFF 가 먼저 위치함.
            If abytTemp(0) = &HFF And abytTemp(1) = &HFE Then
                Call CopyMemory(abytTemp(0), abytTemp(2), lngUpper - 1)
                ReDim Preserve abytTemp(lngUpper - 2)
                lngUpper = lngUpper - 2
            End If
        End If
        lngCharCount = (lngUpper + 1) \ 2

        ' 이제 변환에 필요한 메모리의 크기를 구합니다.
        lngByteNum = WideCharToMultiByteArray(CP_UTF8, 0, abytTemp(0), lngCharCount, 0, 0, 0, 0)

        If lngByteNum > 0 Then
            ReDim abytUTF8(lngByteNum - 1)
            lngByteNum = WideCharToMultiByteArray(CP_UTF8, 0, abytTemp(0), lngCharCount, abytUTF8(0), lngByteNum, 0, 0)
            lngUpper = UBound(abytUTF8)
            ' 변환되어 있는 UTF-8 바이트 배열 선두에 UTF-8 표식을 넣기 위해
            ' 기존의 바이트 배열을 뒤로 밀어내고, 배열 앞부분에 표식을 추가합니다.
            ReDim Preserve abytUTF8(lngUpper + 3)
            Call CopyMemory(abytUTF8(3), abytUTF8(0), lngUpper + 1)
            abytUTF8(0) = &HEF
            abytUTF8(1) = &HBB
            abytUTF8(2) = &HBF

            UTF8FromUTF16withMark = VB6.CopyArray(abytUTF8)
        End If
        Exit Function

ConversionErr:
        MsgBox(" Conversion failed ")

    End Function

    Private Sub MakeLOF()
        Dim intRow As Short
        Dim intCol As Short
        'FIXIT: 'strdata'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                           FixIT90210ae-R1672-R1B8ZE
        Dim strdata As Object
        'FIXIT: 'varTmp'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                            FixIT90210ae-R1672-R1B8ZE
        Dim varTmp As Object
        Dim abytUTF16() As Byte
        Dim abytUTF8() As Byte

        'Cancel을 True로 설정합니다.
        'UPGRADE_WARNING: Visual Basic .NET에서는 CommonDialog CancelError 속성이 지원되지 않습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
        CommonDialog1.CancelError = True

        On Error GoTo ErrHandler

        'Flags 속성을 설정합니다.
        'UPGRADE_WARNING: MSComDlg.CommonDialog 속성 CommonDialog1.Flags이(가) 새로운 동작을 가진 CommonDialog1Font.ShowEffects(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
        CommonDialog1Font.ShowEffects = True
        'UPGRADE_ISSUE: cdlCFBoth 상수가 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
        'UPGRADE_ISSUE: MSComDlg.CommonDialog 속성 CommonDialog1.flags이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        CommonDialog1.Flags = MSComDlg.FontsConstants.cdlCFBoth

        '[글꼴] 대화 상자를 표시합니다.
        CommonDialog1Save.ShowDialog()

        'FIXIT: 'Right' 함수를 'Right$' 함수로 바꾸십시오.                                                    FixIT90210ae-R9757-R1B8ZE
        'FIXIT: 'LCase' 함수를 'LCase$' 함수로 바꾸십시오.                                                    FixIT90210ae-R9757-R1B8ZE
        If Not LCase(VB.Right(CommonDialog1Save.FileName, 4)) = ".lof" Then
            CommonDialog1Save.FileName = CommonDialog1Save.FileName & ".lof"
        End If

        FileOpen(1, CommonDialog1Save.FileName, OpenMode.Binary)
        With spdList
            'UPGRADE_WARNING: strdata 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            strdata = ""
            For intRow = 1 To .MaxRows
                For intCol = 1 To .MaxCols - 1 '-- 마지막 Control제거
                    .GetText(intCol, intRow, varTmp)
                    'UPGRADE_WARNING: varTmp 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: strdata 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    strdata = strdata & varTmp & "^"
                Next
                'UPGRADE_WARNING: strdata 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strdata = strdata & vbCr
            Next

        End With

        'UPGRADE_WARNING: strdata 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        abytUTF16 = strdata
        'abytUTF16 = "유니코드 인코딩 변환 테스트 : UTF-16 LE 를 UTF-8 방식으로 변환하기"
        abytUTF8 = UTF8FromUTF16withMark(abytUTF16)

        'Open "C:\_UTF8TestFile.TXT" For Binary As #1
        'UPGRADE_WARNING: Put이(가) FilePut(으)로 업그레이드되어 새 동작을 가집니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(1, abytUTF8)
        FileClose(1)
        'MsgBox " 변환 완료. " & vbCrLf & " 인터넷 익스플로러로 _UTF8TestFile.TXT 파일을 확인할 수 있습니다. "


        FileClose(1)

        Exit Sub

ErrHandler:

    End Sub

    Private Sub MakeJOB()
        Dim intRow As Short
        Dim intCol As Short
        'FIXIT: 'strdata'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                           FixIT90210ae-R1672-R1B8ZE
        Dim strdata As Object
        'FIXIT: 'varTmp'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                            FixIT90210ae-R1672-R1B8ZE
        Dim varTmp As Object

        On Error GoTo ErrHandler

        FileOpen(1, My.Application.Info.DirectoryPath & "\" & gWork & "Job.txt", OpenMode.Output)

        'FIXIT: Print method 은(는) Visual Basic .NET에서 해당되는 항목이 없으므로 업그레이드되지 않습니다.                  FixIT90210ae-R7593-R67265
        Print(1, "[JobPK]" & Chr(13) & Chr(10))
        'FIXIT: Print method 은(는) Visual Basic .NET에서 해당되는 항목이 없으므로 업그레이드되지 않습니다.                  FixIT90210ae-R7593-R67265
        Print(1, Me.Text & ";" & VB6.Format(Now, "yyyy-mm-dd") & ";A;A;A;1;V" & Chr(13) & Chr(10))

        With spdList
            'FIXIT: Print method 은(는) Visual Basic .NET에서 해당되는 항목이 없으므로 업그레이드되지 않습니다.                  FixIT90210ae-R7593-R67265
            Print(1, "[S_Text]" & Chr(13) & Chr(10))
            '        strData = ""
            '        For intRow = 1 To .MaxRows
            '            .GetText 2, intRow, varTmp
            '            If varTmp = "0" Then
            '                .GetText 3, intRow, varTmp
            '                strData = strData & varTmp & ";"
            '                .GetText 22, intRow, varTmp
            '                strData = strData & varTmp
            '                Print #1, strData & Chr(13) + Chr(10);
            '                strData = ""
            '            End If
            '        Next

            '[D_Text]
            'FIXIT: Print method 은(는) Visual Basic .NET에서 해당되는 항목이 없으므로 업그레이드되지 않습니다.                  FixIT90210ae-R7593-R67265
            Print(1, "[D_Text]" & Chr(13) & Chr(10))
            'UPGRADE_WARNING: strdata 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            strdata = ""
            For intRow = 1 To .MaxRows
                .GetText(2, intRow, varTmp)
                'UPGRADE_WARNING: varTmp 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If varTmp = "1" Then
                    .GetText(3, intRow, varTmp)
                    'UPGRADE_WARNING: varTmp 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: strdata 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    strdata = strdata & varTmp & ";"
                    .GetText(22, intRow, varTmp)
                    'UPGRADE_WARNING: varTmp 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: strdata 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    strdata = strdata & varTmp
                    'FIXIT: Print method 은(는) Visual Basic .NET에서 해당되는 항목이 없으므로 업그레이드되지 않습니다.                  FixIT90210ae-R7593-R67265
                    'UPGRADE_WARNING: strdata 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Print(1, strdata & Chr(13) & Chr(10))
                    'UPGRADE_WARNING: strdata 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    strdata = ""
                End If
            Next

            '[S_Image]
            'FIXIT: Print method 은(는) Visual Basic .NET에서 해당되는 항목이 없으므로 업그레이드되지 않습니다.                  FixIT90210ae-R7593-R67265
            Print(1, "[S_Image]" & Chr(13) & Chr(10))
            'UPGRADE_WARNING: strdata 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            strdata = ""
            For intRow = 1 To .MaxRows
                .GetText(2, intRow, varTmp)
                'UPGRADE_WARNING: varTmp 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If varTmp = "2" Then
                    .GetText(3, intRow, varTmp)
                    'UPGRADE_WARNING: varTmp 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: strdata 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    strdata = strdata & varTmp & ";"
                    .GetText(17, intRow, varTmp)
                    'strData = strData & varTmp
                    'UPGRADE_WARNING: strdata 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    strdata = strdata & "0"
                    'FIXIT: Print method 은(는) Visual Basic .NET에서 해당되는 항목이 없으므로 업그레이드되지 않습니다.                  FixIT90210ae-R7593-R67265
                    'UPGRADE_WARNING: strdata 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Print(1, strdata & Chr(13) & Chr(10))
                    'UPGRADE_WARNING: strdata 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    strdata = ""
                End If
            Next

            '[D_Image]
            'FIXIT: Print method 은(는) Visual Basic .NET에서 해당되는 항목이 없으므로 업그레이드되지 않습니다.                  FixIT90210ae-R7593-R67265
            Print(1, "[D_Image]" & Chr(13) & Chr(10))
            'UPGRADE_WARNING: strdata 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            strdata = ""
            For intRow = 1 To .MaxRows
                .GetText(2, intRow, varTmp)
                'UPGRADE_WARNING: varTmp 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If varTmp = "3" Then
                    .GetText(3, intRow, varTmp)
                    'UPGRADE_WARNING: varTmp 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: strdata 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    strdata = strdata & varTmp & ";"
                    .GetText(17, intRow, varTmp)
                    'UPGRADE_WARNING: varTmp 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    varTmp = Split(varTmp, "\")
                    'UPGRADE_WARNING: varTmp() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: strdata 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    strdata = strdata & varTmp(UBound(varTmp))
                    'FIXIT: Print method 은(는) Visual Basic .NET에서 해당되는 항목이 없으므로 업그레이드되지 않습니다.                  FixIT90210ae-R7593-R67265
                    'UPGRADE_WARNING: strdata 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Print(1, strdata & Chr(13) & Chr(10))
                    'UPGRADE_WARNING: strdata 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    strdata = ""
                End If
            Next

            '[Barcode]
            'FIXIT: Print method 은(는) Visual Basic .NET에서 해당되는 항목이 없으므로 업그레이드되지 않습니다.                  FixIT90210ae-R7593-R67265
            Print(1, "[Barcode]" & Chr(13) & Chr(10))
            'UPGRADE_WARNING: strdata 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            strdata = ""
            For intRow = 1 To .MaxRows
                .GetText(2, intRow, varTmp)
                'UPGRADE_WARNING: varTmp 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If varTmp = "4" Then
                    .GetText(22, intRow, varTmp)
                    'UPGRADE_WARNING: varTmp 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: strdata 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    strdata = strdata & varTmp
                    'FIXIT: Print method 은(는) Visual Basic .NET에서 해당되는 항목이 없으므로 업그레이드되지 않습니다.                  FixIT90210ae-R7593-R67265
                    'UPGRADE_WARNING: strdata 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Print(1, strdata & Chr(13) & Chr(10))
                    'UPGRADE_WARNING: strdata 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    strdata = ""
                End If
            Next

        End With

        FileClose(1)

        MsgBox(Me.Text & "의 작업파일이 생성되었습니다. ", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, Me.Text)

        Exit Sub

ErrHandler:

    End Sub

    Public Sub mnuNew_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNew.Click

        Call FrmInitial()

        'FIXIT: 'sNo1'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                              FixIT90210ae-R1672-R1B8ZE
        Dim sNo1 As Object
        Dim sNo2 As String
        Dim intCnt As Short
        Dim strEditObjName As String
        Dim strWLayout As String
        'Dim strHLayout As String

AgainInput:

        'FIXIT: 'Mid' 함수를 'Mid$' 함수로 바꾸십시오.                                                        FixIT90210ae-R9757-R1B8ZE
        'UPGRADE_WARNING: sNo1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        sNo1 = CDbl(Mid(gLayOutValue(CInt(gLayOutUse)), 1, InStr(gLayOutValue(CInt(gLayOutUse)), ":") - 1)) / 10
        'FIXIT: 'Mid' 함수를 'Mid$' 함수로 바꾸십시오.                                                        FixIT90210ae-R9757-R1B8ZE
        sNo2 = CStr(CDbl(Mid(gLayOutValue(CInt(gLayOutUse)), InStr(gLayOutValue(CInt(gLayOutUse)), ":") + 1)) / 10)

        '    sNo1 = InputBox("라벨용지 높이를 입력하세요 [단위 : cm]", "높이 입력", "7.5")
        '
        '    If Len(sNo1) > 0 Then
        '        If Not IsNumeric(sNo1) Then
        '            MsgBox "숫자만 입력하세요.!", vbCritical
        '            GoTo AgainInput
        '        Else
        '            sNo2 = InputBox("라벨용지 넓이를 입력하세요 [단위 : cm]", "넓이 입력", "3.5")
        '            If Len(sNo2) > 0 Then
        '                If Not IsNumeric(sNo2) Then
        '                    MsgBox "숫자만 입력하세요.!", vbCritical
        '                    GoTo AgainInput
        '                End If
        '
        '            End If
        '        End If
        '    End If


        'UPGRADE_WARNING: sNo1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If sNo1 <> "" And sNo2 <> "" Then
            'UPGRADE_WARNING: sNo1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtPaperHSize.Text = sNo1 '/ 10
            txtPaperWSize.Text = sNo2 '/ 10

            'UPGRADE_WARNING: sNo1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            sNo1 = System.Math.Round(sNo1 * CM_TOTWIP, 0)
            sNo2 = CStr(System.Math.Round(CDbl(sNo2) * CM_TOTWIP, 0))

            sstType.SelectedIndex = 5
            '-- Left
            txtTitle.Text = "LINE_L" '항목명(뷰어)
            txtTag.Text = "LINE_L" '항목명(실제)
            gblCtrlNm = "LINE_L" '항목명(실제)
            txtXpos.Text = "1" 'X 좌표
            txtYpos.Text = "1" 'Y 좌표
            txtLineHSize.Text = "1" '선굵기
            'UPGRADE_WARNING: sNo1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtLineWSize.Text = sNo1 '라인폭
            chkLineRotate.CheckState = CShort("1") '라인회전
            chkPrint.CheckState = CShort("0") '출력여부

            strEditObjName = objMake()
            If strEditObjName = "0" Then
                '객체생성 성공
                Call MakeSpdSaveList(txtTitle, (sstType.SelectedIndex))
            End If

            '-- Right
            txtTitle.Text = "LINE_R" '항목명(뷰어)
            txtTag.Text = "LINE_R" '항목명(실제)
            gblCtrlNm = "LINE_R" '항목명(실제)
            txtXpos.Text = sNo2 'X 좌표
            txtYpos.Text = "1" 'Y 좌표
            txtLineHSize.Text = "1" '선굵기
            'UPGRADE_WARNING: sNo1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtLineWSize.Text = sNo1 '라인폭
            chkLineRotate.CheckState = CShort("1") '라인회전
            chkPrint.CheckState = CShort("0") '출력여부

            strEditObjName = objMake()
            If strEditObjName = "0" Then
                '객체생성 성공
                Call MakeSpdSaveList(txtTitle, (sstType.SelectedIndex))
            End If

            '-- Top
            txtTitle.Text = "LINE_T" '항목명(뷰어)
            txtTag.Text = "LINE_T" '항목명(실제)
            gblCtrlNm = "LINE_T" '항목명(실제)
            txtXpos.Text = "1" 'X 좌표
            txtYpos.Text = "1" 'Y 좌표
            txtLineHSize.Text = "1" '선굵기
            txtLineWSize.Text = sNo2 '라인폭
            chkLineRotate.CheckState = CShort("0") '라인회전
            chkPrint.CheckState = CShort("0") '출력여부

            strEditObjName = objMake()
            If strEditObjName = "0" Then
                '객체생성 성공
                Call MakeSpdSaveList(txtTitle, (sstType.SelectedIndex))
            End If

            '-- Bottom
            txtTitle.Text = "LINE_B" '항목명(뷰어)
            txtTag.Text = "LINE_B" '항목명(실제)
            gblCtrlNm = "LINE_B" '항목명(실제)
            txtXpos.Text = "1" 'X 좌표
            'UPGRADE_WARNING: sNo1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtYpos.Text = sNo1 'Y 좌표
            txtLineHSize.Text = "1" '선굵기
            txtLineWSize.Text = sNo2 '라인폭
            chkLineRotate.CheckState = CShort("0") '라인회전
            chkPrint.CheckState = CShort("0") '출력여부

            strEditObjName = objMake()
            If strEditObjName = "0" Then
                '객체생성 성공
                Call MakeSpdSaveList(txtTitle, (sstType.SelectedIndex))
            End If

        End If

    End Sub

    Public Sub mnuSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSave.Click
        Dim i As Short

        Call MakeLOF()

    End Sub

    Public Sub mnuSet_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSet.Click

        frmConfig.Show()

    End Sub

    Public Sub mnuView_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuView.Click

        'If MsgBox("작업파일을 생성하시겠습니까?", vbYesNo + vbInformation, Me.Caption) = vbYes Then
        Call MakeJOB()

        Call Shell(My.Application.Info.DirectoryPath & "\" & "NOTEPAD.EXE", AppWinStyle.NormalFocus)

        Me.WindowState = System.Windows.Forms.FormWindowState.Minimized

        'End If

    End Sub


    'UPGRADE_WARNING: 폼이 초기화될 때 optDevide.CheckedChanged 이벤트가 발생합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub optDevide_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDevide.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Short = optDevide.GetIndex(eventSender)
            Dim intRow As Short
            Dim intCol As Short
            Dim strBuf() As String

            'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
            gDevide = optDevide(Index).Tag

            ' 컬렉션 초기화
            'UPGRADE_NOTE: m_ColCommandButton 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            m_ColCommandButton = Nothing
            m_ColCommandButton = New Collection

            With spdList
                For intRow = 1 To .MaxRows
                    .Row = intRow
                    .Col = 1
                    Erase strBuf
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    If Trim(.Text) <> "" Then
                        ReDim Preserve strBuf(.MaxCols)
                        For intCol = 2 To .MaxCols
                            .Col = intCol
                            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                            strBuf(intCol - 1) = Trim(.Text)
                        Next
                        Call MakeLayout(strBuf)
                        Erase strBuf
                    End If
                Next
            End With

        End If
    End Sub


    Private Sub picDelobj_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles picDelobj.Click
        Dim intRow As Short
        'FIXIT: 'strObjType'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                        FixIT90210ae-R1672-R1B8ZE
        Dim strObjType As Object
        'FIXIT: 'strObjName'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                        FixIT90210ae-R1672-R1B8ZE
        Dim strObjName As Object
        'FIXIT: 'strObjRotate'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                      FixIT90210ae-R1672-R1B8ZE
        Dim strObjRotate As Object

        CType(Me.Controls(txtTag.Text), Object).Visible = False

        Dim counter As Short
        With spdList
            counter = .MaxRows
            For intRow = 1 To counter
                .Row = intRow
                Call .GetText(2, intRow, strObjType)
                Call .GetText(28, intRow, strObjName)
                '
                'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                'UPGRADE_WARNING: strObjName 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: strObjType 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If strObjType = sstType.SelectedIndex And strObjName = Trim(txtTag.Text) Then
                    .Action = FPSpread.ActionConstants.ActionDeleteRow
                    .MaxRows = .MaxRows - 1
                    Exit For
                End If
            Next
        End With
    End Sub

    Private Sub picFont_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles picFont.Click
        Dim Index As Short = picFont.GetIndex(eventSender)

        'Cancel을 True로 설정합니다.
        'UPGRADE_WARNING: Visual Basic .NET에서는 CommonDialog CancelError 속성이 지원되지 않습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
        CommonDialog1.CancelError = True
        On Error GoTo ErrHandler

        'Flags 속성을 설정합니다.
        'UPGRADE_WARNING: MSComDlg.CommonDialog 속성 CommonDialog1.Flags이(가) 새로운 동작을 가진 CommonDialog1Font.ShowEffects(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
        CommonDialog1Font.ShowEffects = True
        'UPGRADE_ISSUE: cdlCFBoth 상수가 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
        'UPGRADE_ISSUE: MSComDlg.CommonDialog 속성 CommonDialog1.flags이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        CommonDialog1.Flags = MSComDlg.FontsConstants.cdlCFBoth

        '폰트 속성을 설정합니다.[Default]
        CommonDialog1Font.Font = VB6.FontChangeName(CommonDialog1Font.Font, "굴림")
        CommonDialog1Font.Font = VB6.FontChangeSize(CommonDialog1Font.Font, 9)

        '[글꼴] 대화 상자를 표시합니다.
        CommonDialog1Font.ShowDialog()
        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
        txtFontName(Index).Text = CommonDialog1Font.Font.Name
        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
        txtFontSize(Index).Text = CStr(CommonDialog1Font.Font.Size)
        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
        chkFontBold(Index).CheckState = IIf(CommonDialog1Font.Font.Bold = True, 1, 0)
        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
        chkFontItalic(Index).CheckState = IIf(CommonDialog1Font.Font.Italic = True, 1, 0)
        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
        chkFontUnder(Index).CheckState = IIf(CommonDialog1Font.Font.Underline = True, 1, 0)

        Exit Sub

ErrHandler:
        '" 사용자가 [취소] 단추를 눌렀습니다.
        Exit Sub

    End Sub

    Private Sub picImage_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles picImage.Click
        Dim Index As Short = picImage.GetIndex(eventSender)

        Dim sFile As String
        sFile = ShowOpen("JPG파일(*.jpg)|*.jpg", My.Application.Info.DirectoryPath & "\" & gImage)
        If sFile <> "" Then
            'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
            txtImageName(Index).Text = sFile
            If Index = 0 Then
                'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
                Didim_SImg.Image = System.Drawing.Image.FromFile(txtImageName(Index).Text)
                'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
                txtImageWSize(Index).Text = CStr(System.Math.Round(VB6.PixelsToTwipsX(Didim_SImg.Width) / CDbl(gScaleCal), 0))
                'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
                txtImageHSize(Index).Text = CStr(System.Math.Round(VB6.PixelsToTwipsY(Didim_SImg.Height) / CDbl(gScaleCal), 0))

                'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
                txtImageWSize(Index + 2).Text = txtImageWSize(Index).Text
                'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
                txtImageHSize(Index + 2).Text = txtImageHSize(Index).Text

                'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
                txtImageDevide(Index).Focus()
            Else
                'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
                Didim_DImg.Image = System.Drawing.Image.FromFile(txtImageName(Index).Text)
                'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
                txtImageWSize(Index).Text = CStr(System.Math.Round(VB6.PixelsToTwipsX(Didim_DImg.Width) / CDbl(gScaleCal), 0))
                'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
                txtImageHSize(Index).Text = CStr(System.Math.Round(VB6.PixelsToTwipsY(Didim_DImg.Height) / CDbl(gScaleCal), 0))

                'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
                txtImageWSize(Index + 2).Text = txtImageWSize(Index).Text
                'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
                txtImageHSize(Index + 2).Text = txtImageHSize(Index).Text

                'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
                txtImageDevide(Index).Focus()
            End If
        Else
            '        MsgBox "You pressed cancel"
        End If

    End Sub

    Private Sub picMake_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles picMake.Click

        '-- Mode Set [생성]
        intMode = 3

        Call objNewMake()

    End Sub

    Private Sub picPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles picPrint.Click
        Call cmdPrint_Click(cmdPrint, New System.EventArgs())
    End Sub

    Private Sub picSet_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles picSet.Click

        '-- Mode Set [적용가능]
        If intMode = 1 Then
            Call objSet()
        End If

    End Sub

    Private Sub Picture1_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles Picture1.MouseDown
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim x As Single = eventArgs.X
        Dim y As Single = eventArgs.Y
        '
        '    If Button = 1 Then
        '        Picture1.Cls '=============>다시 그리기
        ''        Picture1.CurrentX = X
        ''        Picture1.CurrentY = Y
        '        DrawX = X '=========>눌려진좌표기억
        '        DrawY = Y
        '
        '        Picture1.DrawMode = 10
        '
        '        Ot_X = X
        '        Ot_Y = Y
        '    End If

    End Sub

    Private Sub Picture1_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles Picture1.MouseMove
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim x As Single = eventArgs.X
        Dim y As Single = eventArgs.Y

        '    If Button = 1 Then
        '        Picture1.DrawWidth = 1
        '        Picture1.DrawStyle = 2
        '
        '        Picture1.Line (DrawX, DrawY)-(Ot_X, Ot_Y), vbBlack, B
        '        Picture1.Line (DrawX, DrawY)-(X, Y), vbBlack, B
        '
        '        Ot_X = X
        '        Ot_Y = Y
        '    End If

    End Sub

    Private Sub Picture1_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles Picture1.MouseUp
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim x As Single = eventArgs.X
        Dim y As Single = eventArgs.Y

        '    If Button = 1 Then
        '        Picture1.Line (DrawX, DrawY)-(Ot_X, Ot_Y), vbBlue, B
        '        Picture1.DrawMode = 13
        '        Picture1.DrawWidth = 1
        '        Picture1.DrawStyle = 0 '========>단색(하지 않으면 그대로 점선)
        '        Picture1.Line (DrawX, DrawY)-(X, Y), vbBlue, B
        '    End If

    End Sub

    '-- 컨트롤 초기화
    Private Sub CtrlInitial()

        txtPaperHSize.Text = ""
        txtPaperWSize.Text = ""

        '-- Tab 0
        txtFontName(0).Text = ""
        txtFontSize(0).Text = ""
        chkFontBold(0).CheckState = System.Windows.Forms.CheckState.Unchecked
        chkFontUnder(0).CheckState = System.Windows.Forms.CheckState.Unchecked
        chkFontItalic(0).CheckState = System.Windows.Forms.CheckState.Unchecked
        txtContent(0).Text = ""

        '-- Tab 1
        txtFontName(1).Text = ""
        txtFontSize(1).Text = ""
        chkFontBold(1).CheckState = System.Windows.Forms.CheckState.Unchecked
        chkFontUnder(1).CheckState = System.Windows.Forms.CheckState.Unchecked
        chkFontItalic(1).CheckState = System.Windows.Forms.CheckState.Unchecked
        txtContent(1).Text = ""

        '-- Tab 2
        txtImageName(0).Text = ""
        txtImageWSize(0).Text = ""
        txtImageHSize(0).Text = ""
        txtImageWSize(2).Text = ""
        txtImageHSize(2).Text = ""

        chkIStatic.CheckState = System.Windows.Forms.CheckState.Unchecked

        '-- Tab 3
        txtImageName(1).Text = ""
        txtImageWSize(1).Text = ""
        txtImageHSize(1).Text = ""
        txtImageWSize(3).Text = ""
        txtImageHSize(3).Text = ""

        '-- Tab 4
        txtBarDevide.Text = ""
        txtBarWSize.Text = ""
        txtBarHSize.Text = ""
        txtBarData.Text = ""
        chkBarRotate.CheckState = System.Windows.Forms.CheckState.Unchecked

        '-- Tab 5
        txtLineHSize.Text = ""
        txtLineWSize.Text = ""
        chkLineRotate.CheckState = System.Windows.Forms.CheckState.Unchecked

        gblCtrlNm = ""
        gblCtrlIdx = 0


    End Sub

    '-- 화면 초기화
    Private Sub FrmInitial()
        Dim Printer As New Printer
        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
        Dim x As Printer
        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
        Dim prtSelectPrinter As Printer
        Dim boolPrinter_Select_Fales As Boolean
        Dim strDefault As String
        Dim Buffer As String
        Dim aryPrinter() As String
        Dim strBuffer As String
        Dim i As Short
        Dim j As Short

        ' 클래스 초기화
        ClsEventMonitor = New ClassEventMonitor
        m_ColCommandButton = New Collection

        Call CtrlInitial()

        '구분
        cboType.Items.Clear()
        cboType.Items.Add("S_Text")
        cboType.Items.Add("D_Text")
        cboType.Items.Add("S_Image")
        cboType.Items.Add("D_Image")
        cboType.Items.Add("Barcode")
        cboType.Items.Add("Line")

        cboType.SelectedIndex = 0

        '바코드 타입
        cboBarType.Items.Clear()
        cboBarType.Items.Add("None")
        cboBarType.Items.Add("2of5[공통]") '5
        cboBarType.Items.Add("Interleaved2of5[공통]") '6
        cboBarType.Items.Add("3of9[공통]") '0
        cboBarType.Items.Add("Codabar[공통]") '9
        cboBarType.Items.Add("3of9X[공통]") '1
        cboBarType.Items.Add("Code128A[공통]") '11
        cboBarType.Items.Add("Code128B[공통]") '12
        cboBarType.Items.Add("Code128C[공통]") '13
        cboBarType.Items.Add("UPCA[공통]") '15
        cboBarType.Items.Add("MSI[공통]") '7
        cboBarType.Items.Add("Code93[공통]") '3
        cboBarType.Items.Add("ExtendedCode93[공통]") '4
        cboBarType.Items.Add("EAN13[공통]") '17
        cboBarType.Items.Add("EAN8[공통]") '18
        cboBarType.Items.Add("PostNet[공통]") '23
        cboBarType.Items.Add("ANSI3of9[신규]") '
        cboBarType.Items.Add("ANSI3of9X[신규]") '
        cboBarType.Items.Add("Code128Auto[공통]") '10
        cboBarType.Items.Add("UCCEAN128[공통]") '27
        cboBarType.Items.Add("UPCE[공통]") '16
        cboBarType.Items.Add("RoyalMail[신규]") '
        cboBarType.Items.Add("MSICode2[공통]") '8  ??MSIPlessey
        cboBarType.Items.Add("DUN14[공통]") '28

        cboBarType.SelectedIndex = 7

        ' 0:Code39
        ' 1:Code39Extended
        ' 2:Code39Trioptic  x
        ' 3:Code93
        ' 4:Code93Extended
        ' 5:Code2of5
        ' 6:Interleave2of5
        ' 7:MSICode
        ' 8:MSIPlessey
        ' 9:Codabar
        '10:Code128
        '11:Code128A
        '12:Code128B
        '13:Code128C
        '14:Code11          x
        '15:UPCA
        '16:UPCE
        '17:EAN13
        '18:EAN8
        '19:EAN99           x
        '20:JAN8            x
        '21:JAN13           x
        '22:Telepen         x
        '23:PostNet
        '24:RM4SCC          x
        '25:PZN             x
        '26:ISBN            x
        '27:UCCEAN128       x
        '28:DUN14           x


        With spdList
            .MaxRows = 0
            .MaxCols = 29
            '        .SetText 1, 0, "설정순번"
            '        .SetText 2, 0, "항목구분"
            '        .SetText 3, 0, "항목명"
            '        .SetText 4, 0, "X1좌표"
            '        .SetText 5, 0, "X2좌표"
            '        .SetText 6, 0, "Y1좌표"
            '        .SetText 7, 0, "Y2좌표"
            '        .SetText 8, 0, "폰트명"
            '        .SetText 9, 0, "폰트사이즈"
            '        .SetText 10, 0, "굵기"
            '        .SetText 11, 0, "비틀림"
            '        .SetText 12, 0, "밑줄"
            '        .SetText 13, 0, "폰트회전"
            '        .SetText 14, 0, "바코드종류"
            '        .SetText 15, 0, "바코드폭"
            '        .SetText 16, 0, "바코드회전"
            '        .SetText 17, 0, "이미지경로"
            '        .SetText 18, 0, "라인회전"
            '        .SetText 19, 0, "라인두께"
            '        .SetText 20, 0, "라인폭"
            '        .SetText 21, 0, "출력여부"
            '        .SetText 22, 0, "출력값"
            '        .SetText 23, 0, "X좌표 보정값"
            '        .SetText 24, 0, "Y좌표 보정값"
            '        .SetText 25, 0, "용지높이"
            '        .SetText 26, 0, "용지폭"
            '        .SetText 27, 0, "무조건고정"
            '        .SetText 28, 0, "용지방향"
            '        .SetText 29, 0, "Tag"
            '        .ColWidth(-1) = 10 '10
            '        .ColWidth(29) = 0
        End With

        '-- 프린터
        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
        For Each x In Printers
            'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
            cmbPrinter.Items.Add(x.DeviceName)
        Next x

        strBuffer = Space(1024)

        i = GetProfileString("windows", "Device", "", strBuffer, Len(strBuffer))
        aryPrinter = Split(strBuffer, ",")
        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
        strDefault = Trim(aryPrinter(0))

        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
        For Each prtSelectPrinter In Printers
            j = j + 1
            'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            'FIXIT: 'UCase' 함수를 'UCase$' 함수로 바꾸십시오.                                                    FixIT90210ae-R9757-R1B8ZE
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            'FIXIT: 'UCase' 함수를 'UCase$' 함수로 바꾸십시오.                                                    FixIT90210ae-R9757-R1B8ZE
            If UCase(Trim(prtSelectPrinter.DeviceName)) = UCase(Trim(strDefault)) Then
                'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
                Printer = prtSelectPrinter
                boolPrinter_Select_Fales = True
                cmbPrinter.SelectedIndex = j - 1
                Exit For
            End If
        Next prtSelectPrinter

        '-- 가로
        If optHW(0).Checked = True Then
            txtPaperHSize.Text = ""
            txtPaperWSize.Text = ""

            '-- 세로
        Else

        End If

        '-- Mode Set
        intMode = 0

        '-- 바코드 이미지명 초기화
        strBarImgName = ""

        gOpenFileNm = ""

    End Sub

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Private Sub frmLabelDesign_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
        Dim x As Printer
        'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R5481-H1984
        Dim prtSelectPrinter As Printer
        Dim boolPrinter_Select_Fales As Boolean
        Dim strDefault As String
        Dim Buffer As String
        Dim aryPrinter() As String
        Dim strBuffer As String
        Dim i As Short
        Dim j As Short
        '    Dim strLicense As String
        '    Dim strKey  As String
        '
        '    strLicense = "License"
        '
        '    strKey = GetString(HKEY_CURRENT_USER, REG_POSITION, strLicense)
        '
        '    If strKey = "" Or Not IsDate(strKey) And strKey < Format(Now) Then
        '        MsgBox "라이센스 기간이 만료되었거나 없습니다." & vbNewLine & "개발자에게 문의하십시요", vbCritical + vbOKOnly, Me.Caption
        '        End
        '    End If

        ' 버전 정보 표시
        'FIXIT: App.Revision property 은(는) Visual Basic .NET에서 해당되는 항목이 없으므로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
        Me.Text = Me.Text & " [Ver " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Revision & "]"

        'Combo1.ListIndex = 1

        Call MDIForm_Tool()

        Call FrmInitial()

        Call GetSetup()

        txtDevide.Text = gDevide


        '==== API 파일 오픈 관련 =================================================
        ReDim CustomColors(16 * 4 - 1)

        For i = LBound(CustomColors) To UBound(CustomColors)
            CustomColors(i) = 0
        Next i
        '==== API 파일 오픈 관련 =================================================

        'FIXIT: Visual Basic .NET에서는 런타임에 'ScaleMode'을(를) 변경할 수 없습니다.                              FixIT90210ae-R8024-R57265
        'UPGRADE_ISSUE: Form 속성 frmLabelDesign.ScaleMode는 지원되지 않습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8027179A-CB3B-45C0-9863-FAA1AF983B59"'
        Me.ScaleMode = gScaleMode

        Me.Top = 0
        Me.Left = 0
        'FIXIT: Visual Basic .NET에서는 런타임에 'ScaleWidth'을(를) 변경할 수 없습니다.                             FixIT90210ae-R8024-R57265
        'UPGRADE_ISSUE: Form 속성 frmLabelDesign.ScaleWidth이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        Me.ScaleWidth = 1272
        'FIXIT: Visual Basic .NET에서는 런타임에 'ScaleHeight'을(를) 변경할 수 없습니다.                            FixIT90210ae-R8024-R57265
        'UPGRADE_ISSUE: Form 속성 frmLabelDesign.ScaleHeight이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        Me.ScaleHeight = 890

        '    Picture1.ScaleMode = vbTwips


    End Sub

    Private Function ShowOpen(ByRef Ufilter As String, ByRef Upath As String) As String

        OFName.lStructSize = Len(OFName)
        OFName.hwndOwner = Me.Handle.ToInt32
        OFName.hInstance = VB6.GetHInstance.ToInt32
        OFName.lpstrFilter = Ufilter
        OFName.lpstrFile = Space(254)
        OFName.nMaxFile = 255
        OFName.lpstrFileTitle = Space(254)
        OFName.nMaxFileTitle = 255
        OFName.lpstrInitialDir = Upath
        OFName.lpstrTitle = "Open File"
        OFName.flags = 0

        If GetOpenFileName(OFName) Then
            ShowOpen = Trim(OFName.lpstrFile)
            'ShowOpen = Mid(ShowOpen, 1, Len(ShowOpen) - 1)
        Else
            ShowOpen = ""
        End If

    End Function
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Private Sub frmLabelDesign_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        ' 컬렉션 초기화
        'UPGRADE_NOTE: m_ColCommandButton 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        m_ColCommandButton = Nothing
        'UPGRADE_NOTE: ClsEventMonitor 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        ClsEventMonitor = Nothing

    End Sub

    '문자열의 byte를 되돌려 준다.
    Function LengthByte(ByVal Var As String) As Integer
        Dim Cnt As Integer
        Dim num As Integer
        Dim TMP As String

        Cnt = 0 : num = 0
        If Var = "" Then Exit Function
        Do
            'FIXIT: 'Mid' 함수를 'Mid$' 함수로 바꾸십시오.                                                        FixIT90210ae-R9757-R1B8ZE
            'FIXIT: 'Mid' 함수를 'Mid$' 함수로 바꾸십시오.                                                        FixIT90210ae-R9757-R1B8ZE
            Cnt = Cnt + 1 : TMP = Mid(Var, Cnt, 1) : num = num + 1
            If Asc(TMP) < 0 Then num = num + 1
        Loop Until Cnt >= Len(Var)
        LengthByte = num
    End Function

    '-- 오픈한 LOF 파일을 스프레드에 표시한다,
    '-- 용도 : 적용,저장시 사용한다.
    'FIXIT: 'varBuf'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                            FixIT90210ae-R1672-R1B8ZE
    Private Sub SetList(ByRef varBuf As Object)
        Dim intCnt As Short
        Dim intCol As Short
        Dim intRow As Short

        With spdList
            .MaxRows = .MaxRows + 1
            intRow = .MaxRows
            For intCnt = 0 To UBound(varBuf) '- 1
                If .MaxRows = 1 And intCnt = 0 Then
                    'FIXIT: 'Right' 함수를 'Right$' 함수로 바꾸십시오.                                                    FixIT90210ae-R9757-R1B8ZE
                    'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If Len(varBuf(intCnt)) > 1 Then varBuf(intCnt) = VB.Right(varBuf(intCnt), 1)
                    'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .SetText(intCnt + 1, intRow, CStr(varBuf(intCnt)))
                Else
                    If intCnt = UBound(varBuf) Then
                        'UPGRADE_WARNING: varBuf(1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If varBuf(1) = "4" Then
                            .SetText(intCnt + 1, intRow, strBarImgName)
                        Else
                            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                            .SetText(intCnt + 1, intRow, Trim(txtTag.Text))
                        End If
                    Else
                        'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .SetText(intCnt + 1, intRow, CStr(varBuf(intCnt)))
                    End If
                End If
            Next
            .set_RowHeight(-1, 16)
        End With

    End Sub

    'FIXIT: 'idx'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                               FixIT90210ae-R1672-R1B8ZE
    Private Function BarIdxMapper(ByRef idx As Object) As String


        Select Case idx
            Case 0 : BarIdxMapper = CStr(3)
            Case 1 : BarIdxMapper = CStr(5)
            Case 2 : BarIdxMapper = ""
            Case 3 : BarIdxMapper = CStr(11)
            Case 4 : BarIdxMapper = CStr(12)
            Case 5 : BarIdxMapper = CStr(1)
            Case 6 : BarIdxMapper = CStr(2)
            Case 7 : BarIdxMapper = CStr(10)
            Case 8 : BarIdxMapper = CStr(22)
            Case 9 : BarIdxMapper = CStr(4)
            Case 10 : BarIdxMapper = CStr(18)
            Case 11 : BarIdxMapper = CStr(6)
            Case 12 : BarIdxMapper = CStr(7)
            Case 13 : BarIdxMapper = CStr(8)
            Case 14 : BarIdxMapper = ""
            Case 15 : BarIdxMapper = CStr(9)
            Case 16 : BarIdxMapper = CStr(20)
            Case 17 : BarIdxMapper = CStr(13)
            Case 18 : BarIdxMapper = CStr(14)
            Case 19 : BarIdxMapper = ""
            Case 20 : BarIdxMapper = ""
            Case 21 : BarIdxMapper = ""
            Case 22 : BarIdxMapper = ""
            Case 23 : BarIdxMapper = CStr(15)
            Case 24 : BarIdxMapper = ""
            Case 25 : BarIdxMapper = ""
            Case 26 : BarIdxMapper = ""
            Case 27 : BarIdxMapper = ""
            Case 28 : BarIdxMapper = ""
            Case Else : BarIdxMapper = ""
        End Select



    End Function

    Private Sub PaintLine()
        'FIXIT: 'obj'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                               FixIT90210ae-R1672-R1B8ZE
        Dim obj As Object
        Dim ClsEventObject As ClassEventObject
        Dim i As Short

        '-- 가로라인그리기
        For i = 1 To 100
            'ReMake:
            txtTag.Text = "LineW_" & i
            ClsEventObject = New ClassEventObject
            obj = ClsEventObject.CreateObject_Renamed(Me, ClsEventMonitor, ClassEventMonitor.EventObjectID.EventObjectLine, txtTag.Text)
            If Not obj Is Nothing Then
                'UPGRADE_WARNING: obj.X1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                obj.X1 = 0
                'UPGRADE_WARNING: obj.X2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                obj.X2 = 1000
                'UPGRADE_WARNING: obj.Y1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                obj.Y1 = i * 15
                'UPGRADE_WARNING: obj.Y2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                obj.Y2 = i * 15
                'UPGRADE_WARNING: obj.BorderColor 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                obj.BorderColor = &H8000000F '&HE0E0E0
                'UPGRADE_WARNING: obj.BorderStyle 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                obj.BorderStyle = 1
                'UPGRADE_WARNING: obj.BorderWidth 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                obj.BorderWidth = 1
            Else
                'UPGRADE_NOTE: ClsEventObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                ClsEventObject = Nothing
                'UPGRADE_NOTE: obj 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                obj = Nothing
                '            GoTo ReMake

                Exit Sub
            End If

            'UPGRADE_WARNING: obj.Visible 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            obj.Visible = True
            'UPGRADE_WARNING: obj.Container 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            obj.Container = Picture1
            m_ColCommandButton.Add(ClsEventObject)
            'UPGRADE_NOTE: ClsEventObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            ClsEventObject = Nothing
        Next

        '-- 세로라인그리기
        For i = 1 To 100
            txtTag.Text = "LineH_" & i
            ClsEventObject = New ClassEventObject
            obj = ClsEventObject.CreateObject_Renamed(Me, ClsEventMonitor, ClassEventMonitor.EventObjectID.EventObjectLine, txtTag.Text)
            If Not obj Is Nothing Then
                'UPGRADE_WARNING: obj.X1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                obj.X1 = i * 15
                'UPGRADE_WARNING: obj.X2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                obj.X2 = i * 15
                'UPGRADE_WARNING: obj.Y1 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                obj.Y1 = 0
                'UPGRADE_WARNING: obj.Y2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                obj.Y2 = 1000
                'UPGRADE_WARNING: obj.BorderColor 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                obj.BorderColor = &H8000000F '&HE0E0E0
                'UPGRADE_WARNING: obj.BorderStyle 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                obj.BorderStyle = 1
                'UPGRADE_WARNING: obj.BorderWidth 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                obj.BorderWidth = 1
            Else
                'UPGRADE_NOTE: ClsEventObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                ClsEventObject = Nothing
                Exit Sub
            End If

            'UPGRADE_WARNING: obj.Visible 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            obj.Visible = True
            'UPGRADE_WARNING: obj.Container 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            obj.Container = Picture1
            m_ColCommandButton.Add(ClsEventObject)
            'UPGRADE_NOTE: ClsEventObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            ClsEventObject = Nothing
        Next

    End Sub

    '-- 구분별로 오프젝트 내역을 각 항목에 표시한다.
    '   구분[varBuf(1)] 0:SText,1:DText,2:SImage,3:DImage,4:Barcode,5:Line
    'FIXIT: 'varBuf'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                            FixIT90210ae-R1672-R1B8ZE
    Private Sub MakeLayout(ByRef varBuf As Object)
        Dim strEditObjName As String
        Dim i As Short
        Dim strFVar As String
        'FIXIT: 'strTmp'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                            FixIT90210ae-R1672-R1B8ZE
        Dim strTmp As Object

MakeAgain:

        'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        sstType.SelectedIndex = varBuf(1)

        'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        txtPaperHSize.Text = varBuf(25)
        'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        txtPaperWSize.Text = varBuf(25)

        strFVar = ""
        For i = 1 To Len(varBuf(0))
            'FIXIT: 'Mid' 함수를 'Mid$' 함수로 바꾸십시오.                                                        FixIT90210ae-R9757-R1B8ZE
            'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Asc(Mid(varBuf(0), i, 1)) <> 63 Then
                'FIXIT: 'Mid' 함수를 'Mid$' 함수로 바꾸십시오.                                                        FixIT90210ae-R9757-R1B8ZE
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strFVar = strFVar & Mid(varBuf(0), i, 1)
            Else
                'Stop
            End If
        Next

        Select Case varBuf(1)
            Case 0 '## Static Label ##
                'txtTag.Text = Replace(varBuf(2), "-", "_")          '항목명(실제)
                txtTag.Text = "Control_" & strFVar
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtTitle.Text = varBuf(2) '항목명(뷰어)
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtXpos.Text = varBuf(3) 'X 좌표
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtYpos.Text = varBuf(5) 'Y 좌표
                'UPGRADE_WARNING: varBuf(7) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtFontName(0).Text = varBuf(7) '폰트명
                'UPGRADE_WARNING: varBuf(8) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtFontSize(0).Text = varBuf(8) '폰트크기
                'UPGRADE_WARNING: varBuf(9) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                chkFontBold(0).CheckState = varBuf(9) '    굵게
                'UPGRADE_WARNING: varBuf(11) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                chkFontUnder(0).CheckState = varBuf(11) '    밑줄
                'UPGRADE_WARNING: varBuf(10) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                chkFontItalic(0).CheckState = varBuf(10) '    기울게
                'UPGRADE_WARNING: varBuf(21) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtContent(0).Text = varBuf(21) 'Text
                'txtContent1.Text = varBuf(21)                     'Text
                '            txtContent(0).Font.Charset = 163
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                chkTStatic.CheckState = varBuf(26) '무조건고정
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                chkPrint.CheckState = IIf(varBuf(20) = "1", "0", "1") '출력안함

            Case 1 '## Dynamic Label ##
                'txtTag.Text = Replace(varBuf(2), "-", "_")          '항목명(실제)
                txtTag.Text = "Control_" & strFVar
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtTitle.Text = varBuf(2) '항목명(뷰어)
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtXpos.Text = varBuf(3) 'X 좌표
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtYpos.Text = varBuf(5) 'Y 좌표
                'UPGRADE_WARNING: varBuf(7) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtFontName(1).Text = varBuf(7) '폰트명
                'UPGRADE_WARNING: varBuf(8) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtFontSize(1).Text = varBuf(8) '폰트크기
                'UPGRADE_WARNING: varBuf(9) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                chkFontBold(1).CheckState = varBuf(9) '    굵게
                'UPGRADE_WARNING: varBuf(11) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                chkFontUnder(1).CheckState = varBuf(11) '    밑줄
                'UPGRADE_WARNING: varBuf(10) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                chkFontItalic(1).CheckState = varBuf(10) '    기울게
                'UPGRADE_WARNING: varBuf(21) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtContent(1).Text = varBuf(21) 'Text
                '            txtContent(1).Font.Charset = ""
                '            txtContent(1).Font.Charset = 163
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                chkPrint.CheckState = IIf(varBuf(20) = "1", "0", "1") '출력안함

            Case 2 '## Static Image ##
                'txtTag.Text = Replace(varBuf(2), "-", "_")          '항목명(실제)
                txtTag.Text = "Control_" & strFVar
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtTitle.Text = varBuf(2) '항목명(뷰어)
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtXpos.Text = varBuf(3) 'X 좌표
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtYpos.Text = varBuf(5) 'Y 좌표
                'UPGRADE_WARNING: varBuf(16) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtImageName(0).Text = varBuf(16) '이미지경로
                'UPGRADE_WARNING: varBuf(4) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtImageWSize(0).Text = varBuf(4) '      가로SIZE
                'UPGRADE_WARNING: varBuf(6) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtImageHSize(0).Text = varBuf(6) '      세로SIZE
                'UPGRADE_WARNING: varBuf(4) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtImageWSize(2).Text = varBuf(4) '      가로SIZE
                'UPGRADE_WARNING: varBuf(6) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtImageHSize(2).Text = varBuf(6) '      세로SIZE

                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                chkIStatic.CheckState = varBuf(26) '무조건고정
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                chkPrint.CheckState = IIf(varBuf(20) = "1", "0", "1") '출력안함

            Case 3 '## Dynamic Image ##
                'txtTag.Text = Replace(varBuf(2), "-", "_")          '항목명(실제)
                txtTag.Text = "Control_" & strFVar
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtTitle.Text = varBuf(2) '항목명(뷰어)
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtXpos.Text = varBuf(3) 'X 좌표
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtYpos.Text = varBuf(5) 'Y 좌표
                'UPGRADE_WARNING: varBuf(16) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtImageName(1).Text = varBuf(16) '이미지경로
                'UPGRADE_WARNING: varBuf(4) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtImageWSize(1).Text = varBuf(4) '      가로SIZE
                'UPGRADE_WARNING: varBuf(6) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtImageHSize(1).Text = varBuf(6) '      세로SIZE
                'UPGRADE_WARNING: varBuf(4) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtImageWSize(3).Text = varBuf(4) '      가로SIZE
                'UPGRADE_WARNING: varBuf(6) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtImageHSize(3).Text = varBuf(6) '      세로SIZE
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                chkPrint.CheckState = IIf(varBuf(20) = "1", "0", "1") '출력안함

            Case 4
                'txtTag.Text = Replace(varBuf(2), "-", "_")          '항목명(실제)
                txtTag.Text = "Control_" & strFVar
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtTitle.Text = varBuf(2) '항목명(뷰어)


                '-- 바코드 타입 기존 프로그램과 신규프로그램 Mapping
                'UPGRADE_WARNING: strTmp 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strTmp = BarIdxMapper(varBuf(13))
                'UPGRADE_WARNING: strTmp 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If strTmp = "" Then
                    cboBarType.SelectedIndex = 7 '바코드 타입
                Else
                    'UPGRADE_WARNING: strTmp 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    cboBarType.SelectedIndex = strTmp '바코드 타입
                End If

                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtXpos.Text = varBuf(3) 'X 좌표
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtYpos.Text = varBuf(5) 'Y 좌표
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtBarData.Text = varBuf(21) '바코드Data
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtBarWSize.Text = varBuf(4) '      길이SIZE
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtBarHSize.Text = varBuf(6) '      세로SIZE
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                chkBarRotate.CheckState = IIf(varBuf(15) = "2", "1", "0") '     회전
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                chkPrint.CheckState = IIf(varBuf(20) = "1", "0", "1") '출력안함

            Case 5
                'txtTag.Text = Replace(varBuf(2), "-", "_")          '항목명(실제)
                txtTag.Text = "Control_" & strFVar
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtTitle.Text = varBuf(2) '항목명(뷰어)
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtXpos.Text = varBuf(3) 'X 좌표
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtYpos.Text = varBuf(5) 'Y 좌표
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                chkLineRotate.CheckState = IIf(varBuf(17) = "0", "0", "1") '라인회전
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtLineHSize.Text = varBuf(18) '선굵기
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtLineWSize.Text = varBuf(19) '라인폭
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                chkPrint.CheckState = IIf(varBuf(20) = "1", "0", "1") '출력안함
        End Select

        '-- 객체이름 업데이트
        gblCtrlNm = txtTag.Text
        gblCtrlIdx = CShort(strFVar)

        '-- 객체생성
        strEditObjName = objMake()

        If strEditObjName = "0" Then
            '객체생성 성공
        Else
            '객체생성 실패
            'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            varBuf(2) = strEditObjName
            GoTo MakeAgain
        End If

    End Sub


    Private Sub SetLayout(ByRef intTabidx As Short)

        '구분[varBuf(1)] 0:SText,1:DText,2:SImage,3:DImage,4:Barcode,5:Line

        Dim intCnt As Short
        Dim intCol As Short
        Dim intRow As Short
        Dim strIdx As String
        Dim strTitle As String

        With spdList
            For intRow = 1 To .MaxRows
                '항목구분,항목명 비교
                .Row = intRow
                'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                .Col = 2 : strIdx = Trim(.Text)
                'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                .Col = 29 : strTitle = Trim(.Text)
                '            If findSameCtrlNm(3, txtTitle.Text) Then
                '                MsgBox "동일한 항목명은 사용할 수 없습니다.", vbInformation, Me.Caption
                '                Exit For
                '            End If
                'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                If intTabidx = CDbl(strIdx) And Trim(txtTag.Text) = Trim(strTitle) Then
                    Select Case intTabidx
                        Case 0
                            .SetText(3, intRow, txtTitle.Text)
                            .SetText(4, intRow, txtXpos.Text)
                            .SetText(6, intRow, txtYpos.Text)
                            .SetText(8, intRow, txtFontName(0).Text)
                            .SetText(9, intRow, txtFontSize(0).Text)
                            .SetText(10, intRow, IIf(chkFontBold(0).CheckState = CDbl("0"), "0", "1"))
                            .SetText(11, intRow, IIf(chkFontItalic(0).CheckState = CDbl("0"), "0", "1"))
                            .SetText(12, intRow, IIf(chkFontUnder(0).CheckState = CDbl("0"), "0", "1"))
                            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                            .SetText(22, intRow, Trim(txtContent(0).Text))
                            .SetText(21, intRow, IIf(chkPrint.CheckState = CDbl("1"), "0", "1")) '출력여부
                            .SetText(27, intRow, IIf(chkTStatic.CheckState = CDbl("0"), "0", "1")) '무조건고정

                        Case 1
                            .SetText(3, intRow, txtTitle.Text)
                            .SetText(4, intRow, txtXpos.Text)
                            .SetText(6, intRow, txtYpos.Text)
                            .SetText(8, intRow, txtFontName(1).Text)
                            .SetText(9, intRow, txtFontSize(1).Text)
                            .SetText(10, intRow, IIf(chkFontBold(1).CheckState = CDbl("0"), "0", "1"))
                            .SetText(11, intRow, IIf(chkFontItalic(1).CheckState = CDbl("0"), "0", "1"))
                            .SetText(12, intRow, IIf(chkFontUnder(1).CheckState = CDbl("0"), "0", "1"))
                            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                            .SetText(22, intRow, Trim(txtContent(1).Text))
                            .SetText(21, intRow, IIf(chkPrint.CheckState = CDbl("1"), "0", "1")) '출력여부

                        Case 2
                            .SetText(3, intRow, txtTitle.Text)
                            .SetText(4, intRow, txtXpos.Text)
                            .SetText(5, intRow, txtImageWSize(0).Text)
                            .SetText(6, intRow, txtYpos.Text)
                            .SetText(7, intRow, txtImageHSize(0).Text)
                            .SetText(17, intRow, txtImageName(0).Text)

                            .SetText(21, intRow, IIf(chkPrint.CheckState = CDbl("1"), "0", "1")) '출력여부
                            .SetText(27, intRow, IIf(chkIStatic.CheckState = CDbl("0"), "0", "1")) '무조건고정

                        Case 3
                            .SetText(3, intRow, txtTitle.Text)
                            .SetText(4, intRow, txtXpos.Text)
                            .SetText(5, intRow, txtImageWSize(1).Text)
                            .SetText(6, intRow, txtYpos.Text)
                            .SetText(7, intRow, txtImageHSize(1).Text)
                            .SetText(17, intRow, txtImageName(1).Text)

                            .SetText(21, intRow, IIf(chkPrint.CheckState = CDbl("1"), "0", "1")) '출력여부

                        Case 4
                            .SetText(3, intRow, txtTitle.Text)
                            .SetText(4, intRow, txtXpos.Text)
                            .SetText(5, intRow, txtBarWSize.Text)
                            .SetText(6, intRow, txtYpos.Text)
                            .SetText(7, intRow, txtBarHSize.Text)
                            .SetText(14, intRow, cboBarType.SelectedIndex) '-- 바코드 종류
                            '.SetText 15, intRow, cboBarType.ListIndex    '-- 바코드 폭
                            .SetText(16, intRow, IIf(chkBarRotate.CheckState = CDbl("0"), "0", "2")) '-- 바코드 회전
                            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                            .SetText(22, intRow, Trim(txtBarData.Text)) '-- 바코드 출력값

                            .SetText(21, intRow, IIf(chkPrint.CheckState = CDbl("1"), "0", "1")) '출력여부

                        Case 5
                            .SetText(3, intRow, txtTitle.Text)
                            .SetText(4, intRow, txtXpos.Text)
                            .SetText(5, intRow, txtXpos.Text)
                            .SetText(6, intRow, txtYpos.Text)
                            .SetText(7, intRow, txtLineWSize.Text)
                            .SetText(9, intRow, txtLineHSize.Text)
                            .SetText(18, intRow, IIf(chkLineRotate.CheckState = CDbl("0"), "0", "1")) '라인회전
                            .SetText(19, intRow, txtLineHSize.Text) '라인두께
                            .SetText(20, intRow, txtLineWSize.Text) '라인폭

                            .SetText(21, intRow, IIf(chkPrint.CheckState = CDbl("1"), "0", "1")) '출력여부

                    End Select

                    Exit Sub
                End If
            Next
        End With

    End Sub


    Public Function toUTF8(ByVal szSource As String) As String
        On Error GoTo ErrHandler

        Dim szChar As String
        Dim WideChar As Integer
        Dim nLength As Short
        Dim i As Short

        nLength = Len(szSource)

        For i = 1 To nLength
            'FIXIT: 'Mid' 함수를 'Mid$' 함수로 바꾸십시오.                                                        FixIT90210ae-R9757-R1B8ZE
            szChar = Mid(szSource, i, 1)

            If Asc(szChar) < 0 Then
                'FIXIT: 키워드 'MidB'은(는) Visual Basic .NET에서 지원되지 않습니다.                                      FixIT90210ae-R6614-H1984
                'FIXIT: 키워드 'MidB'은(는) Visual Basic .NET에서 지원되지 않습니다.                                      FixIT90210ae-R6614-H1984
                'UPGRADE_ISSUE: MidB 함수는 지원되지 않습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
                'UPGRADE_ISSUE: AscB 함수는 지원되지 않습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
                WideChar = CInt(AscB(MidB(szChar, 2, 1))) * 256 + AscB(MidB(szChar, 1, 1))

                If (WideChar And &HFF80) = 0 Then
                    'FIXIT: 'Hex' 함수를 'Hex$' 함수로 바꾸십시오.                                                        FixIT90210ae-R9757-R1B8ZE
                    toUTF8 = toUTF8 & Hex(WideChar)
                ElseIf (WideChar And &HF000) = 0 Then
                    'FIXIT: 'Hex' 함수를 'Hex$' 함수로 바꾸십시오.                                                        FixIT90210ae-R9757-R1B8ZE
                    'FIXIT: 'Hex' 함수를 'Hex$' 함수로 바꾸십시오.                                                        FixIT90210ae-R9757-R1B8ZE
                    toUTF8 = toUTF8 & Hex(CShort(CShort(WideChar And &HFFC0) / 64) Or &HC0) & Hex(WideChar And &H3F Or &H80)
                Else
                    'FIXIT: 'Hex' 함수를 'Hex$' 함수로 바꾸십시오.                                                        FixIT90210ae-R9757-R1B8ZE
                    'FIXIT: 'Hex' 함수를 'Hex$' 함수로 바꾸십시오.                                                        FixIT90210ae-R9757-R1B8ZE
                    'FIXIT: 'Hex' 함수를 'Hex$' 함수로 바꾸십시오.                                                        FixIT90210ae-R9757-R1B8ZE
                    toUTF8 = toUTF8 & Hex(CShort(CShort(WideChar And &HF000) / 4096) Or &HE0) & Hex(CShort(CShort(WideChar And &HFFC0) / 64) And &H3F Or &H80) & Hex(WideChar And &H3F Or &H80)

                End If
            Else
                'FIXIT: 'Hex' 함수를 'Hex$' 함수로 바꾸십시오.                                                        FixIT90210ae-R9757-R1B8ZE
                toUTF8 = toUTF8 & Hex(Asc(szChar))
            End If
        Next

        Exit Function

ErrHandler:
        toUTF8 = ""

    End Function

    Public Function URLEncode(ByRef URLStr As String) As String

        Dim sURL As String '** 입력받은 URL 문자열
        Dim sBuffer As String '** URL 인코딩 처리 중 URL 을 담을 버퍼 문자열
        Dim sTemp As String '** 임시 문자열
        'UPGRADE_NOTE: cChar이(가) cChar_Renamed(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Dim cChar_Renamed As String '** URL 문자열 중 현재 인텍스의 문자
        Dim lErrNum As Integer '** 오류 번호
        Dim sErrSource As String '** 오류 소스
        Dim sErrDesc As String '** 소류 설명
        Dim sMsg As String '** 오류 메세지
        Dim Index As Short

        On Error GoTo ErrorHanddle

        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
        sURL = Trim(URLStr) '** URL 문자열을 얻는다.
        sBuffer = "" '** 임시 버퍼용 문자열 변수 초기화.

        '******************************************************
        '* URL 인코딩 작업
        '******************************************************

        For Index = 1 To Len(sURL)
            '** 현재 인덱스의 문자를 얻는다.
            'FIXIT: 'Mid' 함수를 'Mid$' 함수로 바꾸십시오.                                                        FixIT90210ae-R9757-R1B8ZE
            cChar_Renamed = Mid(sURL, Index, 1)

            If cChar_Renamed = "0" Or (cChar_Renamed >= "1" And cChar_Renamed <= "9") Or (cChar_Renamed >= "a" And cChar_Renamed <= "z") Or (cChar_Renamed >= "A" And cChar_Renamed <= "Z") Or cChar_Renamed = "-" Or cChar_Renamed = "_" Or cChar_Renamed = "." Or cChar_Renamed = "*" Then
                '** URL 에 허용되는 문자들 :: 버퍼 문자열에 추가한다.
                sBuffer = sBuffer & cChar_Renamed
            ElseIf cChar_Renamed = " " Then
                '** 공백 문자 :: + 로 대체하여 버퍼 문자열에 추가한다.
                sBuffer = sBuffer & "+"
            Else
                '** URL 에 허용되지 않는 문자들 :: % 로 인코딩해서 버퍼 문자열에 추가한다.
                'FIXIT: 'Hex' 함수를 'Hex$' 함수로 바꾸십시오.                                                        FixIT90210ae-R9757-R1B8ZE
                sTemp = CStr(Hex(Asc(cChar_Renamed)))
                If Len(sTemp) = 4 Then
                    'FIXIT: 'Mid' 함수를 'Mid$' 함수로 바꾸십시오.                                                        FixIT90210ae-R9757-R1B8ZE
                    'FIXIT: 'Left' 함수를 'Left$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    sBuffer = sBuffer & "%" & VB.Left(sTemp, 2) & "%" & Mid(sTemp, 3, 2)
                ElseIf Len(sTemp) = 2 Then
                    sBuffer = sBuffer & "%" & sTemp
                End If
            End If
        Next

        '** 결과를 리턴한다.
        URLEncode = sBuffer

        Exit Function

ErrorHanddle:

        '** 오류가 발생하면 공백 문자를 리턴한다.
        URLEncode = ""

        '** 오류 정보를 얻는다.
        lErrNum = Err.Number
        sErrSource = Err.Source
        sErrDesc = Err.Description

        '** 이벤트 로그에 오류를 기록한다.
        sMsg = vbCrLf & vbCrLf & "Error Object : EgoCube.URLTools," & vbCrLf & "Error Method : Public Function URLEncode(URLStr As String) As String," & vbCrLf & "Error Number : " & lErrNum & "," & vbCrLf & "Error Source : " & sErrSource & "," & vbCrLf & "Error Description : " & sErrDesc

        'FIXIT: App.LogEvent method 은(는) Visual Basic .NET에서 해당되는 항목이 없으므로 업그레이드되지 않습니다.           FixIT90210ae-R7593-R67265
        My.Application.Log.WriteEntry(sMsg, System.Diagnostics.TraceEventType.Error)

        '** 오류를 발생시킨다.
        Err.Raise(lErrNum, sErrSource, sErrDesc)


        Exit Function


    End Function

    Public Sub mnuOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuOpen.Click
        'FIXIT: 'strSrcfile'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                        FixIT90210ae-R1672-R1B8ZE
        Dim strSrcfile As Object
        'FIXIT: 'varBuffer'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                         FixIT90210ae-R1672-R1B8ZE
        Dim varBuffer() As Object
        'FIXIT: 'varBuf'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                            FixIT90210ae-R1672-R1B8ZE
        Dim varBuf As Object
        Dim lngBufLen As Integer
        Dim i As Integer
        'FIXIT: 'Buffer'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                            FixIT90210ae-R1672-R1B8ZE
        Dim Buffer As Object
        Dim BufChar As String
        Dim j As Integer
        Dim bytBuff() As Byte

        Static ChkSumCnt As Integer
        Dim strTxt As String

        Dim FileNumber As Integer
        Dim FileName As String
        Dim FileCount As Integer
        Dim LineCount As Integer
        Dim FileOpenNumber As Short
        Dim data As String
        Dim splitdata() As String

        Dim utf8() As Byte
        'FIXIT: 'ucs2'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                              FixIT90210ae-R1672-R1B8ZE
        Dim ucs2 As Object
        Dim chars As Integer
        'FIXIT: 'varTmp'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                            FixIT90210ae-R1672-R1B8ZE
        Dim varTmp As Object

        ' 폼초기화
        Call FrmInitial()

        ''    'Cancel을 True로 설정합니다.
        ''    CommonDialog1.CancelError = True
        ''    On Error GoTo ErrHandler
        ''
        ''    '경로 속성을 설정합니다.
        ''    CommonDialog1.InitDir = App.Path & "\" & gLayOut
        ''    CommonDialog1.Filter = "LayoutFile(*.lof)|*.lof"
        ''
        ''    '[파일] 대화 상자를 표시합니다.
        ''    CommonDialog1.ShowOpen
        ''    strSrcfile = CommonDialog1.FileName
        ''
        ''    '컬렉션 초기화
        ''    Set m_ColCommandButton = Nothing
        ''    Set m_ColCommandButton = New Collection
        ''
        ''    'LOF 파일 열기
        ''    FileName = CommonDialog1.FileName


        Dim sFile As String
        FileName = ShowOpen("LayoutFile(*.lof)|*.lof", My.Application.Info.DirectoryPath & "\" & gLayOut)
        If FileName <> "" Then
            'UPGRADE_WARNING: varTmp 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            varTmp = Split(FileName, "\")
            'UPGRADE_WARNING: varTmp() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Me.Text = varTmp(UBound(varTmp))
            FileOpenNumber = FreeFile()
            LineCount = 0

            '====== 유니코드 테스트
            '''    Dim strBuffer
            '''    Dim dY As Long
            '''
            '''    dY = 1
            '''
            '''    Open FileName For Input As #3
            '''
            '''    strBuffer = ""
            '''    Do While Not EOF(3)
            '''        textbox = textbox & Input(1, #3)
            '''    Loop
            '''
            '''textbox = Mid(textbox, 1000)
            '''    Close #3
            '''
            ''''    Debug.Print strBuffer
            '''
            '''    Picture1.FontName = textbox.Font
            '''    'Picture1.Font = "Calibri"
            ''''    textbox.Text = ucs2
            '''    Call TextOutW(Picture1.hdc, 10, dY * 50, StrPtr(textbox), Len(textbox))
            '''Exit Sub
            '====== 유니코드 테스트43


            gOpenFileNm = FileName

            FileOpen(1, FileName, OpenMode.Binary) 'UTF-8 문서지정
            ReDim utf8(LOF(1))

            'UPGRADE_WARNING: Get이(가) FileGet(으)로 업그레이드되어 새 동작을 가집니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            FileGet(1, utf8)

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


            '''    chars = MultiByteToWideChar(CP_UTF8, 0, VarPtr(utf8(0)), LOF(1), 0, 0)
            '''    ucs2 = Space(chars)
            '''    chars = MultiByteToWideChar(CP_UTF8, 0, VarPtr(utf8(0)), LOF(1), StrPtr(ucs2), chars)
            '''    varBuf = Split(ucs2, Chr(13))


            'FIXIT: 키워드 'VarPtr'은(는) Visual Basic .NET에서 지원되지 않습니다.                                    FixIT90210ae-R6614-H1984
            'UPGRADE_ISSUE: VarPtr 함수는 지원되지 않습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
            chars = MultiByteToWideChar(CP_UTF8, 0, VarPtr(utf8(0)), LOF(1), 0, 0)
            'UPGRADE_WARNING: ucs2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ucs2 = Space(chars)
            'FIXIT: 키워드 'StrPtr'은(는) Visual Basic .NET에서 지원되지 않습니다.                                    FixIT90210ae-R6614-H1984
            'FIXIT: 키워드 'VarPtr'은(는) Visual Basic .NET에서 지원되지 않습니다.                                    FixIT90210ae-R6614-H1984
            'UPGRADE_ISSUE: StrPtr 함수는 지원되지 않습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
            'UPGRADE_ISSUE: VarPtr 함수는 지원되지 않습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
            chars = MultiByteToWideChar(CP_UTF8, 0, VarPtr(utf8(0)), LOF(1), StrPtr(ucs2), chars)
            'UPGRADE_WARNING: ucs2 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: varBuf 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            varBuf = Split(ucs2, Chr(13))


            '    chars = MultiByteToWideChar(CP_UTF8, 0, VarPtr(utf8(0)), LOF(1), StrPtr(ucs2), chars)

            'textbox.Font.Charset = 163 '베트남어
            '    Call Shell(App.Path & "\" & "NOTEPAD.EXE " & gOpenFileNm, vbNormalFocus)


            '    RichTextBox1 = ucs2
            '    textbox = ucs2
            FileClose(1)

            'Exit Sub


            '오픈한 LOF파일 버퍼에 쓰기
            For i = 0 To UBound(varBuf)
                ReDim Preserve varBuffer(i)
                'UPGRADE_WARNING: varBuf() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: varBuffer(LineCount) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                varBuffer(LineCount) = varBuf(i)
                LineCount = LineCount + 1
            Next


            '오픈한 LOF파일 화면그리기/스프레드쓰기
            For i = 0 To UBound(varBuffer) - 1
                'UPGRADE_WARNING: varBuffer(i) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If varBuffer(i) <> "" Then
                    'UPGRADE_WARNING: varBuffer() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: varBuf 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    varBuf = Split(varBuffer(i), "^")
                    'Debug.Print varBuffer(i)
                    Call MakeLayout(varBuf)
                    Call SetList(varBuf)
                End If
            Next

            Call PaintLine()

            '    intMode = 1
        Else
            '        MsgBox "You pressed cancel"
        End If

        Exit Sub

ErrHandler:

    End Sub



    Private Sub picUndo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles picUndo.Click
        'FIXIT: 'Moveobj'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                           FixIT90210ae-R1672-R1B8ZE
        Dim Moveobj As Object
        'FIXIT: 'x'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                                 FixIT90210ae-R1672-R1B8ZE
        Dim x As Object
        Dim y As Integer

        'UPGRADE_WARNING: LMousePos.obj 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Moveobj 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Moveobj = LMousePos.obj
        'UPGRADE_WARNING: x 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        x = LMousePos.fromx
        y = LMousePos.fromy

        'UPGRADE_WARNING: x 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CType(Me.Controls(Moveobj), Object).Left = VB6.TwipsToPixelsX(x)
        CType(Me.Controls(Moveobj), Object).Top = VB6.TwipsToPixelsY(y)
    End Sub

    Private Sub spdList_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpread._DSpreadEvents_ClickEvent) Handles spdList.ClickEvent

        Call SetControl(eventArgs.row)

    End Sub

    Private Sub SetControl(ByRef intRow As Integer)

        Dim strTmp As String

        With spdList
            .Row = intRow
            '-- 제목
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            .Col = 2 : sstType.SelectedIndex = Trim(.Text)
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            .Col = 3 : txtTitle.Text = Trim(.Text)
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            .Col = 29 : txtTag.Text = Trim(.Text)
            '-- 위치
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            .Col = 4 : txtXpos.Text = Trim(.Text)
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            .Col = 6 : txtYpos.Text = Trim(.Text)
            '-- 넓이,높이(두께)
            Select Case sstType.SelectedIndex
                'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                Case 2 : .Col = 5 : txtImageWSize(0).Text = Trim(.Text)
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    .Col = 7 : txtImageHSize(0).Text = Trim(.Text)
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    .Col = 5 : txtImageWSize(2).Text = Trim(.Text)
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    .Col = 7 : txtImageHSize(2).Text = Trim(.Text)
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                Case 3 : .Col = 5 : txtImageWSize(1).Text = Trim(.Text)
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    .Col = 7 : txtImageHSize(1).Text = Trim(.Text)
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    .Col = 5 : txtImageWSize(3).Text = Trim(.Text)
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    .Col = 7 : txtImageHSize(3).Text = Trim(.Text)
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                Case 4 : .Col = 5 : txtBarWSize.Text = Trim(.Text)
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    .Col = 7 : txtBarHSize.Text = Trim(.Text)
            End Select
            '-- 폰트
            Select Case sstType.SelectedIndex
                'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                Case 0 : .Col = 8 : txtFontName(0).Text = Trim(.Text)
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    .Col = 9 : txtFontSize(0).Text = Trim(.Text)
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    .Col = 10 : chkFontBold(0).CheckState = IIf(Trim(.Text) = "0", "0", "1") '폰트굵게
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    .Col = 11 : chkFontUnder(0).CheckState = IIf(Trim(.Text) = "0", "0", "1") '폰트밑줄
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    .Col = 12 : chkFontItalic(0).CheckState = IIf(Trim(.Text) = "0", "0", "1") '폰트기울게
                    '.Col = 13: chkFontItalic(0).Value = IIf(Trim(.Text) = "0", "0", "1") '폰트회전
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                Case 1 : .Col = 8 : txtFontName(1).Text = Trim(.Text)
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    .Col = 9 : txtFontSize(1).Text = Trim(.Text)
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    .Col = 10 : chkFontBold(1).CheckState = IIf(Trim(.Text) = "0", "0", "1") '폰트굵게
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    .Col = 11 : chkFontUnder(1).CheckState = IIf(Trim(.Text) = "0", "0", "1") '폰트밑줄
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    .Col = 12 : chkFontItalic(1).CheckState = IIf(Trim(.Text) = "0", "0", "1") '폰트기울게
                    '.Col = 13: chkFontItalic(0).Value = IIf(Trim(.Text) = "0", "0", "1") '폰트회전
            End Select
            '-- 바코드
            '-- 바코드 타입 기존 프로그램과 신규프로그램 Mapping
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            .Col = 14 : strTmp = BarIdxMapper(Trim(.Text))
            If strTmp = "" Then
                cboBarType.SelectedIndex = 7
            Else
                cboBarType.SelectedIndex = strTmp
            End If
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            .Col = 15 : txtBarDevide.Text = Trim(.Text)
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            .Col = 16 : chkBarRotate.CheckState = IIf(Trim(.Text) = "0", 0, 2)
            '-- 이미지
            If sstType.SelectedIndex = 3 Then
                'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                .Col = 17 : txtImageName(0).Text = Trim(.Text)
            ElseIf sstType.SelectedIndex = 4 Then
                'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                .Col = 17 : txtImageName(1).Text = Trim(.Text)
            End If
            '-- 라인
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            .Col = 18 : chkLineRotate.CheckState = IIf(Trim(.Text) = "0", 0, 1)
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            .Col = 19 : txtLineHSize.Text = Trim(.Text)
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            .Col = 20 : txtLineWSize.Text = Trim(.Text)
            '-- 출력여부
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            .Col = 21 : chkPrint.CheckState = IIf(Trim(.Text) = "1", 0, 1)
            '-- 출력값
            Select Case sstType.SelectedIndex
                'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                Case 0 : .Col = 22 : txtContent(0).Text = Trim(.Text)
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                Case 1 : .Col = 22 : txtContent(1).Text = Trim(.Text)
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                    'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                Case 4 : .Col = 22 : txtBarData.Text = Trim(.Text)
            End Select
            '-- 무조건고정
            If sstType.SelectedIndex = 0 Then
                'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                .Col = 27 : chkTStatic.CheckState = IIf(Trim(.Text) = "0", 0, 1)
            ElseIf sstType.SelectedIndex = 2 Then
                'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                .Col = 27 : chkIStatic.CheckState = IIf(Trim(.Text) = "0", 0, 1)
            End If

        End With

    End Sub


    Private Sub spdList_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpread._DSpreadEvents_KeyPressEvent) Handles spdList.KeyPressEvent
        'FIXIT: 'varTmp'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                            FixIT90210ae-R1672-R1B8ZE
        Dim varTmp As Object

        If eventArgs.keyAscii = 13 Then

            Call SetControl((spdList.ActiveRow))

            intMode = 1

            Call cmdSet_Click(cmdSet, New System.EventArgs())

        End If

    End Sub

    Private Sub spdList_LeaveRow(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpread._DSpreadEvents_LeaveRowEvent) Handles spdList.LeaveRow

        Call SetControl(eventArgs.newRow)

    End Sub

    'Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '    If X > (Command2.Width - 100) And Y > (Command2.Height - 100) And Button = vbLeftButton Then
    '        drageMode = True
    '    Else
    '        drageMode = False
    '    End If
    '    If drageMode Then
    '        Command2.Height = Y
    '        Command2.Width = X
    '    End If
    'End Sub


    Private Sub sstType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles sstType.SelectedIndexChanged
        Static PreviousTab As Short = sstType.SelectedIndex()
        Select Case sstType.SelectedIndex
            Case 0
                txtTitle.Text = "S_TEXT" & gblCtrlIdx
                'cmdFont(0).SetFocus
            Case 1
                txtTitle.Text = "D_TEXT" & gblCtrlIdx
                'cmdFont(1).SetFocus
            Case 2
                txtTitle.Text = "S_Image" & gblCtrlIdx
                'cmdImage(0).SetFocus
            Case 3
                txtTitle.Text = "D_Image" & gblCtrlIdx
                'cmdImage(1).SetFocus
            Case 4
                txtTitle.Text = "BARCODE" & gblCtrlIdx
                'cboBarType.SetFocus
            Case 5
                txtTitle.Text = "LINE" & gblCtrlIdx
                'txtLineHSize.SetFocus
                txtLineHSize.Text = "1"
        End Select

        txtTag.Text = ""
        txtXpos.Text = CStr(10)
        txtYpos.Text = CStr(10)

        cboType.SelectedIndex = sstType.SelectedIndex

        PreviousTab = sstType.SelectedIndex()
    End Sub



    Private Sub tlbMain_ButtonClick(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles tlbMain.ItemClicked
        Dim Button As System.Windows.Forms.ToolStripItem = CType(eventSender, System.Windows.Forms.ToolStripItem)
        Select Case Button.Name
            Case TLBKEY_NEW
                Call mnuNew_Click(mnuNew, New System.EventArgs())
            Case TLBKEY_OPEN
                Call mnuOpen_Click(mnuOpen, New System.EventArgs())
            Case TLBKEY_SAVE
                Call mnuSave_Click(mnuSave, New System.EventArgs())
            Case TLBKEY_MAKE
                Call mnuMake_Click(mnuMake, New System.EventArgs())
            Case TLBKEY_VIEW
                Call mnuView_Click(mnuView, New System.EventArgs())
            Case TLBKEY_EDIT
                Call mnuSet_Click(mnuSet, New System.EventArgs())
            Case TLBKEY_EDIT
                Call mnuSet_Click(mnuSet, New System.EventArgs())
            Case TLBKEY_EXIT
                Call mnuClose_Click(mnuClose, New System.EventArgs())
        End Select

    End Sub

    Private Sub tmrMove_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrMove.Tick

        Call objMove(intMoveIdx)

    End Sub


    Private Sub txtBarHSize_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtBarHSize.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        If KeyAscii = 13 Then
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            If Not IsNumeric(Trim(txtBarHSize.Text)) Then
                MsgBox("숫자만 입력이 가능합니다.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, Me.Text)
                txtBarHSize.Focus()
            End If
        End If

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtBarWSize_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtBarWSize.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        If KeyAscii = 13 Then
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            If Not IsNumeric(Trim(txtBarWSize.Text)) Then
                MsgBox("숫자만 입력이 가능합니다.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, Me.Text)
                txtBarWSize.Focus()
            End If
        End If

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtDevide_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDevide.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Dim intRow As Short
        Dim intCol As Short
        Dim strBuf() As String

        If KeyAscii = 13 Then
            If IsNumeric(txtDevide.Text) Then
                gDevide = txtDevide.Text

                ' 컬렉션 초기화
                'UPGRADE_NOTE: m_ColCommandButton 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                m_ColCommandButton = Nothing
                m_ColCommandButton = New Collection

                With spdList
                    For intRow = 1 To .MaxRows
                        .Row = intRow
                        .Col = 1
                        Erase strBuf
                        'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                        If Trim(.Text) <> "" Then
                            ReDim Preserve strBuf(.MaxCols)
                            For intCol = 2 To .MaxCols
                                .Col = intCol
                                'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
                                strBuf(intCol - 1) = Trim(.Text)
                            Next
                            Call MakeLayout(strBuf)
                            Erase strBuf
                        End If
                    Next
                End With
            Else
                MsgBox("숫자만 입력이 가능합니다.", MsgBoxStyle.Information, Me.Text)
                txtDevide.Focus()
                GoTo EventExitSub
            End If
        End If
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub


    Private Sub txtFontSize_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFontSize.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Dim Index As Short = txtFontSize.GetIndex(eventSender)

        If KeyAscii = 13 Then
            'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            If Not IsNumeric(Trim(txtFontSize(Index).Text)) Then
                MsgBox("숫자만 입력이 가능합니다.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, Me.Text)
                'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
                txtFontSize(Index).Focus()
            End If
        End If

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub


    Private Sub txtImageDevide_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtImageDevide.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Dim Index As Short = txtImageDevide.GetIndex(eventSender)

        If KeyAscii = 13 Then
            Call cmdImageDevSet_Click(cmdImageDevSet.Item(Index), New System.EventArgs())
        End If

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtImageHSize_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtImageHSize.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Dim Index As Short = txtImageHSize.GetIndex(eventSender)

        If KeyAscii = 13 Then
            'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            If Not IsNumeric(Trim(txtImageHSize(Index).Text)) Then
                MsgBox("숫자만 입력이 가능합니다.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, Me.Text)
                'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
                txtImageHSize(Index).Focus()
            End If
        End If

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtImageWSize_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtImageWSize.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Dim Index As Short = txtImageWSize.GetIndex(eventSender)

        If KeyAscii = 13 Then
            'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            If Not IsNumeric(Trim(txtImageWSize(Index).Text)) Then
                MsgBox("숫자만 입력이 가능합니다.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, Me.Text)
                'FIXIT: Printer 개체 및 Printers 컬렉션은 업그레이드 마법사를 통해 Visual Basic .NET으로 업그레이드되지 않습니다.         FixIT90210ae-R7593-R67265
                txtImageWSize(Index).Focus()
            End If
        End If

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtLineHSize_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtLineHSize.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        If KeyAscii = 13 Then
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            If Not IsNumeric(Trim(txtLineHSize.Text)) Then
                MsgBox("숫자만 입력이 가능합니다.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, Me.Text)
                txtLineHSize.Focus()
            End If
        End If

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtLineWSize_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtLineWSize.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        If KeyAscii = 13 Then
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            If Not IsNumeric(Trim(txtLineWSize.Text)) Then
                MsgBox("숫자만 입력이 가능합니다.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, Me.Text)
                txtLineWSize.Focus()
            End If
        End If

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtPaperHSize_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPaperHSize.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        If KeyAscii = 13 Then
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            If Not IsNumeric(Trim(txtPaperHSize.Text)) Then
                MsgBox("숫자만 입력이 가능합니다.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, Me.Text)
                txtPaperHSize.Focus()
            End If
        End If

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtPaperWSize_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPaperWSize.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        If KeyAscii = 13 Then
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            If Not IsNumeric(Trim(txtPaperWSize.Text)) Then
                MsgBox("숫자만 입력이 가능합니다.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, Me.Text)
                txtPaperWSize.Focus()
            End If
        End If

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    'UPGRADE_WARNING: 폼이 초기화될 때 txtXpos.TextChanged 이벤트가 발생합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub txtXpos_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtXpos.TextChanged

        txtXmm.Text = CStr(CDbl(txtXpos.Text) / 3.779)

    End Sub

    Private Sub txtXpos_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtXpos.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        If KeyAscii = 13 Then
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            If Not IsNumeric(Trim(txtXpos.Text)) Then
                MsgBox("숫자만 입력이 가능합니다.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, Me.Text)
                txtXpos.Focus()
            End If
        End If

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    'UPGRADE_WARNING: 폼이 초기화될 때 txtYpos.TextChanged 이벤트가 발생합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub txtYpos_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtYpos.TextChanged

        txtYmm.Text = CStr(CDbl(txtYpos.Text) / 3.779)

    End Sub

    Private Sub txtYpos_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtYpos.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        If KeyAscii = 13 Then
            'FIXIT: 'Trim' 함수를 'Trim$' 함수로 바꾸십시오.                                                      FixIT90210ae-R9757-R1B8ZE
            If Not IsNumeric(Trim(txtYpos.Text)) Then
                MsgBox("숫자만 입력이 가능합니다.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, Me.Text)
                txtYpos.Focus()
            End If
        End If

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
End Class