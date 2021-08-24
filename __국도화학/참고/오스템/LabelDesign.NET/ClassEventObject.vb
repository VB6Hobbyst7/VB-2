Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class ClassEventObject
	'===============================================================================
	'  프로그램 : 컨트롤 동적 생성 및 이벤트 처리
	'  파 일 명 : ClassEventObject.cls
	'  설    명 :
	'
	'  수정이력
	'===============================================================================
	'  Flag    수정일자    수정자   수정내용
	'-------------------------------------------------------------------------------
	'  [CYJ#0] 2007.03.30
	'  [OSW]   2011.09.21  오세원   라벨,이미지,라인,바코드[Mabry.BarCod] 추가
	'===============================================================================
	
	Private m_ClsEventMonitor As ClassEventMonitor ' 이벤트 전달을 위한 클래스
	Private m_FrmOwner As frmLabelDesign ' 부모 폼 정보
	Private m_IntEventObjectId As ClassEventMonitor.EventObjectID ' CommandButton, TextBox, Label, Image ...
	'FIXIT: 'm_VarParam'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                        FixIT90210ae-R1672-R1B8ZE
	Private m_VarParam() As Object ' 이벤트 발생에서의 파라미터 정보
	
	' 이벤트 처리를 위한 Object
	Private WithEvents EventCommandButton As System.Windows.Forms.Button
	Private WithEvents EventTextBox As System.Windows.Forms.TextBox
	Private WithEvents EventSLabel As System.Windows.Forms.Label
	Private WithEvents EventDLabel As System.Windows.Forms.Label
	Private WithEvents EventBLabel As System.Windows.Forms.Label
	Private WithEvents EventSImage As System.Windows.Forms.PictureBox
	Private WithEvents EventDImage As System.Windows.Forms.PictureBox
	Private WithEvents EventLImage As System.Windows.Forms.PictureBox
	Private WithEvents EventBImage As System.Windows.Forms.PictureBox
	Private WithEvents EventLine As Microsoft.VisualBasic.PowerPacks.LineShape
	Private WithEvents EventBarcode As AxBarcodLib.AxBarcod
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' 클래스 생성자
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'UPGRADE_NOTE: Class_Initialize이(가) Class_Initialize_Renamed(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		' 클래스 초기화
		'UPGRADE_NOTE: EventCommandButton 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EventCommandButton = Nothing
		'UPGRADE_NOTE: EventTextBox 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EventTextBox = Nothing
		'UPGRADE_NOTE: EventSLabel 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EventSLabel = Nothing
		'UPGRADE_NOTE: EventDLabel 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EventDLabel = Nothing
		'UPGRADE_NOTE: EventBLabel 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EventBLabel = Nothing
		'UPGRADE_NOTE: EventSImage 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EventSImage.Image = Nothing
		'UPGRADE_NOTE: EventDImage 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EventDImage.Image = Nothing
		'UPGRADE_NOTE: EventLImage 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EventLImage.Image = Nothing
		'UPGRADE_NOTE: EventBImage 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EventBImage.Image = Nothing
		'UPGRADE_NOTE: EventLine 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EventLine = Nothing
		'UPGRADE_NOTE: EventBarcode 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EventBarcode = Nothing
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' 클래스 소멸자
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'UPGRADE_NOTE: Class_Terminate이(가) Class_Terminate_Renamed(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		
		On Error Resume Next
		
		' 동적 생성 컨트롤 제거
		'UPGRADE_WARNING: Me.EventObject.Name 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_FrmOwner.Controls.Remove(Me.EventObject.Name)
		
		'UPGRADE_NOTE: EventCommandButton 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EventCommandButton = Nothing
		'UPGRADE_NOTE: EventTextBox 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EventTextBox = Nothing
		'UPGRADE_NOTE: EventSLabel 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EventSLabel = Nothing
		'UPGRADE_NOTE: EventDLabel 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EventDLabel = Nothing
		'UPGRADE_NOTE: EventBLabel 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EventBLabel = Nothing
		'UPGRADE_NOTE: EventSImage 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EventSImage.Image = Nothing
		'UPGRADE_NOTE: EventDImage 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EventDImage.Image = Nothing
		'UPGRADE_NOTE: EventLImage 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EventLImage.Image = Nothing
		'UPGRADE_NOTE: EventBImage 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EventBImage.Image = Nothing
		'UPGRADE_NOTE: EventLine 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EventLine = Nothing
		'UPGRADE_NOTE: EventBarcode 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EventBarcode = Nothing
		
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' EventMonitor Property
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	'FIXIT: 'EventMonitor'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                      FixIT90210ae-R1672-R1B8ZE
	Public Property EventMonitor() As Object
		Get
			EventMonitor = m_ClsEventMonitor
		End Get
		Set(ByVal Value As Object)
			m_ClsEventMonitor = Value
		End Set
	End Property
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Owner Property
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	'FIXIT: 'Owner'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                             FixIT90210ae-R1672-R1B8ZE
	Public Property Owner() As Object
		Get
			Owner = m_FrmOwner
		End Get
		Set(ByVal Value As Object)
			m_FrmOwner = Value
		End Set
	End Property
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' EventObject Property
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'FIXIT: 'EventObject'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                       FixIT90210ae-R1672-R1B8ZE
	Public ReadOnly Property EventObject() As Object
		Get
			
			Select Case m_IntEventObjectId
				Case ClassEventMonitor.EventObjectID.EventObjectCommandButton
					EventObject = EventCommandButton
					
				Case ClassEventMonitor.EventObjectID.EventObjectTextBox
					EventObject = EventTextBox
					
				Case ClassEventMonitor.EventObjectID.EventObjectSLabel
					EventObject = EventSLabel
					
				Case ClassEventMonitor.EventObjectID.EventObjectDLabel
					EventObject = EventDLabel
					
				Case ClassEventMonitor.EventObjectID.EventObjectBLabel
					EventObject = EventBLabel
					
				Case ClassEventMonitor.EventObjectID.EventObjectSImage
					EventObject = EventSImage
					
				Case ClassEventMonitor.EventObjectID.EventObjectDImage
					EventObject = EventDImage
					
				Case ClassEventMonitor.EventObjectID.EventObjectLImage
					EventObject = EventLImage
					
				Case ClassEventMonitor.EventObjectID.EventObjectBImage
					EventObject = EventBImage
					
				Case ClassEventMonitor.EventObjectID.EventObjectLine
					EventObject = EventLine
					
				Case ClassEventMonitor.EventObjectID.EventObjectBarcode
					EventObject = EventBarcode
					
			End Select
			
		End Get
	End Property
	
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Param Property
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'FIXIT: 'Param'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                             FixIT90210ae-R1672-R1B8ZE
	Public ReadOnly Property Param(ByVal IntIndex As Short) As Object
		Get
			
			On Error Resume Next
			'UPGRADE_WARNING: m_VarParam(IntIndex) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Param = m_VarParam(IntIndex)
			
		End Get
	End Property
	
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Param Property
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'FIXIT: 'Params'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                            FixIT90210ae-R1672-R1B8ZE
	Public ReadOnly Property Params() As Object
		Get
			
			'UPGRADE_WARNING: Params 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Params = VB6.CopyArray(m_VarParam)
			
		End Get
	End Property
	
	
	'===============================================================================
	' 함 수 명 : CreateObject()
	' 설    명 : 동적 컨트롤 생성
	' 입 력 값 :
	' 결 과 값 :
	'===============================================================================
	'FIXIT: 'CreateObject'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                      FixIT90210ae-R1672-R1B8ZE
	'UPGRADE_NOTE: CreateObject이(가) CreateObject_Renamed(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Public Function CreateObject_Renamed(ByRef FrmOwner As System.Windows.Forms.Form, ByRef ClsEventMonitor As ClassEventMonitor, _
                                         ByVal IntEventObjectId As ClassEventMonitor.EventObjectID, ByVal StrObjectName As String) As Object
        'FIXIT: 'obj'을(를) 초기에 바인딩되는 데이터 형식으로 선언하십시오.                                               FixIT90210ae-R1672-R1B8ZE
        Dim obj As Object
        '    Dim objB            As BarcodLib.Barcod

        On Error Resume Next

        m_FrmOwner = FrmOwner
        m_ClsEventMonitor = ClsEventMonitor
        m_IntEventObjectId = IntEventObjectId

        Dim instance As Control.ControlCollection
        Dim value As Control

        Dim CtrlButton As New System.Windows.Forms.Button
        Dim CtrlText As New System.Windows.Forms.TextBox
        Dim CtrlLabel As New System.Windows.Forms.Label
        Dim CtrlPicture As New System.Windows.Forms.PictureBox

        CtrlButton.Name = StrObjectName

        Select Case IntEventObjectId

            ' CommandButton
            Case ClassEventMonitor.EventObjectID.EventObjectCommandButton
                'UPGRADE_ISSUE: Controls 메서드 m_FrmOwner.Controls.Add이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                m_FrmOwner.Picture1.Controls.Add(CtrlButton)
                EventCommandButton = obj

                ' TextBox
            Case ClassEventMonitor.EventObjectID.EventObjectTextBox
                'UPGRADE_ISSUE: Controls 메서드 m_FrmOwner.Controls.Add이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                m_FrmOwner.Picture1.Controls.Add(CtrlText)
                EventTextBox = obj

                ' Static_Label
            Case ClassEventMonitor.EventObjectID.EventObjectSLabel
                'UPGRADE_ISSUE: Controls 메서드 m_FrmOwner.Controls.Add이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                m_FrmOwner.Picture1.Controls.Add(CtrlLabel)
                EventSLabel = obj

                ' Dynamic_Label
            Case ClassEventMonitor.EventObjectID.EventObjectDLabel
                'UPGRADE_ISSUE: Controls 메서드 m_FrmOwner.Controls.Add이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                m_FrmOwner.Picture1.Controls.Add(CtrlLabel)
                EventDLabel = obj

                ' Barcode_Label
            Case ClassEventMonitor.EventObjectID.EventObjectBLabel
                'UPGRADE_ISSUE: Controls 메서드 m_FrmOwner.Controls.Add이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                m_FrmOwner.Picture1.Controls.Add(CtrlLabel)
                EventBLabel = obj

                ' Static_Image
            Case ClassEventMonitor.EventObjectID.EventObjectSImage
                'UPGRADE_ISSUE: Controls 메서드 m_FrmOwner.Controls.Add이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                m_FrmOwner.Picture1.Controls.Add(CtrlPicture)
                EventSImage = obj

                ' Dynamic_Image
            Case ClassEventMonitor.EventObjectID.EventObjectDImage
                'UPGRADE_ISSUE: Controls 메서드 m_FrmOwner.Controls.Add이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                m_FrmOwner.Picture1.Controls.Add(CtrlPicture)
                EventDImage = obj

                ' Line_Image
            Case ClassEventMonitor.EventObjectID.EventObjectLImage
                'UPGRADE_ISSUE: Controls 메서드 m_FrmOwner.Controls.Add이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                m_FrmOwner.Picture1.Controls.Add(CtrlPicture)
                EventLImage = obj

                ' Barcode_Image
            Case ClassEventMonitor.EventObjectID.EventObjectBImage
                'UPGRADE_ISSUE: Controls 메서드 m_FrmOwner.Controls.Add이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                m_FrmOwner.Picture1.Controls.Add(CtrlPicture)
                EventBImage = obj

                ' Line
            Case ClassEventMonitor.EventObjectID.EventObjectLine
                'UPGRADE_ISSUE: Controls 메서드 m_FrmOwner.Controls.Add이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                m_FrmOwner.Picture1.Controls.Add(CtrlPicture)
                EventLine = obj

                ' Barcode
            Case ClassEventMonitor.EventObjectID.EventObjectBarcode
                'UPGRADE_ISSUE: Controls 메서드 m_FrmOwner.Controls.Add이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                m_FrmOwner.Picture1.Controls.Add(CtrlPicture)
                EventBarcode = obj

            Case Else
                MsgBox("EventObjectId Error!!", MsgBoxStyle.Critical)
                Exit Function

        End Select

        CreateObject_Renamed = obj

    End Function
	
	'===============================================================================
	' 함 수 명 : PfRaiseEvent()
	' 설    명 : 이벤트 발생 시킴
	' 입 력 값 :
	' 결 과 값 :
	' 작 성 일 : 2007.03.30
	' 작 성 자 : 제용재(182cm@korea.com)
	'===============================================================================
	'UPGRADE_WARNING: Params ParamArray가 ByRef에서 ByVal로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93C6A0DC-8C99-429A-8696-35FC4DCEFCCC"'
	Private Sub PfRaiseEvent(ByVal StrEventName As String, ParamArray ByVal Params() As Object)
		
		' 파라미터 저장
		m_VarParam = VB6.CopyArray(Params)
		
		' CommandButton의 Click 이벤트 전달
		Call m_ClsEventMonitor.RaiseUserEvent(Me, StrEventName)
		
	End Sub
	
	
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' CommandButton Event
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	Private Sub EventCommandButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EventCommandButton.Click
		Call PfRaiseEvent("Click")
	End Sub
	
	Private Sub EventCommandButton_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EventCommandButton.Enter
		Call PfRaiseEvent("GotFocus")
	End Sub
	
	Private Sub EventCommandButton_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles EventCommandButton.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Call PfRaiseEvent("KeyDown", KeyCode, Shift)
	End Sub
	
	Private Sub EventCommandButton_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles EventCommandButton.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Call PfRaiseEvent("KeyPress", KeyAscii)
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub EventCommandButton_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles EventCommandButton.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Call PfRaiseEvent("KeyUp", KeyCode, Shift)
	End Sub
	
	Private Sub EventCommandButton_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EventCommandButton.Leave
		Call PfRaiseEvent("LostFocus")
	End Sub
	
	Private Sub EventCommandButton_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles EventCommandButton.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call PfRaiseEvent("MouseDown", Button, Shift, x, y)
	End Sub
	
	Private Sub EventCommandButton_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles EventCommandButton.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call PfRaiseEvent("MouseMove", Button, Shift, x, y)
	End Sub
	
	Private Sub EventCommandButton_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles EventCommandButton.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call PfRaiseEvent("MouseUp", Button, Shift, x, y)
	End Sub
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' TextBox Event
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'UPGRADE_WARNING: 폼이 초기화될 때 EventTextBox.TextChanged 이벤트가 발생합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub EventTextBox_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EventTextBox.TextChanged
		Call PfRaiseEvent("Change")
	End Sub
	
	Private Sub EventTextBox_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EventTextBox.Click
		Call PfRaiseEvent("Click")
	End Sub
	
	Private Sub EventTextBox_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EventTextBox.DoubleClick
		Call PfRaiseEvent("DblClick")
	End Sub
	
	Private Sub EventTextBox_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EventTextBox.Enter
		Call PfRaiseEvent("GotFocus")
	End Sub
	
	Private Sub EventTextBox_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles EventTextBox.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Call PfRaiseEvent("KeyDown", KeyCode, Shift)
	End Sub
	
	Private Sub EventTextBox_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles EventTextBox.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Call PfRaiseEvent("KeyPress", KeyAscii)
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub EventTextBox_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles EventTextBox.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Call PfRaiseEvent("KeyUp", KeyCode, Shift)
	End Sub
	
	Private Sub EventTextBox_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EventTextBox.Leave
		Call PfRaiseEvent("LostFocus")
	End Sub
	
	Private Sub EventTextBox_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles EventTextBox.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call PfRaiseEvent("MouseDown", Button, Shift, x, y)
	End Sub
	
	Private Sub EventTextBox_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles EventTextBox.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call PfRaiseEvent("MouseMove", Button, Shift, x, y)
	End Sub
	
	Private Sub EventTextBox_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles EventTextBox.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call PfRaiseEvent("MouseUp", Button, Shift, x, y)
	End Sub
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Static Label Event
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'FIXIT: Private Sub EventSLabel_Change event 은(는) Visual Basic .NET에서 해당되는 항목이 없으므로 업그레이드되지 않습니다.     FixIT90210ae-R7593-R67265
	'UPGRADE_ISSUE: Label 이벤트 EventSLabel.Change이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub EventSLabel_Change()
		Call PfRaiseEvent("Change")
	End Sub
	
	Private Sub EventSLabel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EventSLabel.Click
		Call PfRaiseEvent("Click")
		Call obj_Click(EventSLabel, 0)
	End Sub
	
	Private Sub EventSLabel_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EventSLabel.DoubleClick
		Call PfRaiseEvent("DblClick")
	End Sub
	
	Private Sub EventSLabel_GotFocus()
		Call PfRaiseEvent("GotFocus")
	End Sub
	
	Private Sub EventSLabel_KeyDown(ByRef KeyCode As Short, ByRef Shift As Short)
		Call PfRaiseEvent("KeyDown", KeyCode, Shift)
	End Sub
	
	Private Sub EventSLabel_KeyPress(ByRef KeyAscii As Short)
		Call PfRaiseEvent("KeyPress", KeyAscii)
	End Sub
	
	Private Sub EventSLabel_KeyUp(ByRef KeyCode As Short, ByRef Shift As Short)
		Call PfRaiseEvent("KeyUp", KeyCode, Shift)
	End Sub
	
	Private Sub EventSLabel_LostFocus()
		Call PfRaiseEvent("LostFocus")
	End Sub
	
	Private Sub EventSLabel_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles EventSLabel.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call PfRaiseEvent("MouseDown", Button, Shift, x, y)
		Call obj_MouseDown(EventSLabel, Button, Shift, x, y)
	End Sub
	
	Private Sub EventSLabel_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles EventSLabel.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call PfRaiseEvent("MouseMove", Button, Shift, x, y)
		Call obj_MouseMove(EventSLabel, Button, Shift, x, y)
	End Sub
	
	Private Sub EventSLabel_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles EventSLabel.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call PfRaiseEvent("MouseUp", Button, Shift, x, y)
	End Sub
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Dynamic Label Event
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'FIXIT: Private Sub EventDLabel_Change event 은(는) Visual Basic .NET에서 해당되는 항목이 없으므로 업그레이드되지 않습니다.     FixIT90210ae-R7593-R67265
	'UPGRADE_ISSUE: Label 이벤트 EventDLabel.Change이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub EventDLabel_Change()
		Call PfRaiseEvent("Change")
	End Sub
	
	Private Sub EventDLabel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EventDLabel.Click
		Call PfRaiseEvent("Click")
		Call obj_Click(EventDLabel, 1)
	End Sub
	
	Private Sub EventDLabel_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EventDLabel.DoubleClick
		Call PfRaiseEvent("DblClick")
	End Sub
	
	Private Sub EventDLabel_GotFocus()
		Call PfRaiseEvent("GotFocus")
	End Sub
	
	Private Sub EventDLabel_KeyDown(ByRef KeyCode As Short, ByRef Shift As Short)
		Call PfRaiseEvent("KeyDown", KeyCode, Shift)
	End Sub
	
	Private Sub EventDLabel_KeyPress(ByRef KeyAscii As Short)
		Call PfRaiseEvent("KeyPress", KeyAscii)
	End Sub
	
	Private Sub EventDLabel_KeyUp(ByRef KeyCode As Short, ByRef Shift As Short)
		Call PfRaiseEvent("KeyUp", KeyCode, Shift)
	End Sub
	
	Private Sub EventDLabel_LostFocus()
		Call PfRaiseEvent("LostFocus")
	End Sub
	
	Private Sub EventDLabel_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles EventDLabel.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call PfRaiseEvent("MouseDown", Button, Shift, x, y)
		Call obj_MouseDown(EventDLabel, Button, Shift, x, y)
	End Sub
	
	Private Sub EventDLabel_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles EventDLabel.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call PfRaiseEvent("MouseMove", Button, Shift, x, y)
		Call obj_MouseMove(EventDLabel, Button, Shift, x, y)
	End Sub
	
	Private Sub EventDLabel_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles EventDLabel.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call PfRaiseEvent("MouseUp", Button, Shift, x, y)
	End Sub
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Barcode Label Event
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'FIXIT: Private Sub EventBLabel_Change event 은(는) Visual Basic .NET에서 해당되는 항목이 없으므로 업그레이드되지 않습니다.     FixIT90210ae-R7593-R67265
	'UPGRADE_ISSUE: Label 이벤트 EventBLabel.Change이(가) 업그레이드되지 않았습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub EventBLabel_Change()
		Call PfRaiseEvent("Change")
	End Sub
	
	Private Sub EventBLabel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EventBLabel.Click
		Call PfRaiseEvent("Click")
		Call obj_Click(EventBLabel, 4)
	End Sub
	
	Private Sub EventBLabel_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EventBLabel.DoubleClick
		Call PfRaiseEvent("DblClick")
	End Sub
	
	Private Sub EventBLabel_GotFocus()
		Call PfRaiseEvent("GotFocus")
	End Sub
	
	Private Sub EventBLabel_KeyDown(ByRef KeyCode As Short, ByRef Shift As Short)
		Call PfRaiseEvent("KeyDown", KeyCode, Shift)
	End Sub
	
	Private Sub EventBLabel_KeyPress(ByRef KeyAscii As Short)
		Call PfRaiseEvent("KeyPress", KeyAscii)
	End Sub
	
	Private Sub EventBLabel_KeyUp(ByRef KeyCode As Short, ByRef Shift As Short)
		Call PfRaiseEvent("KeyUp", KeyCode, Shift)
	End Sub
	
	Private Sub EventBLabel_LostFocus()
		Call PfRaiseEvent("LostFocus")
	End Sub
	
	Private Sub EventBLabel_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles EventBLabel.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call PfRaiseEvent(EventBLabel.Text, "MouseDown", Button, Shift, x, y)
		Call obj_MouseDown(EventBLabel, Button, Shift, x, y)
	End Sub
	
	Private Sub EventBLabel_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles EventBLabel.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call PfRaiseEvent("MouseMove", Button, Shift, x, y)
		Call obj_MouseMove(EventBLabel, Button, Shift, x, y)
	End Sub
	
	Private Sub EventBLabel_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles EventBLabel.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call PfRaiseEvent("MouseUp", Button, Shift, x, y)
	End Sub
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Static Image Event
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	Private Sub EventSImage_Change()
		Call PfRaiseEvent("Change")
	End Sub
	
	Private Sub EventSImage_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EventSImage.Click
		Call PfRaiseEvent("Click")
		Call obj_Click(EventSImage, 2)
	End Sub
	
	Private Sub EventSImage_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EventSImage.DoubleClick
		Call PfRaiseEvent("DblClick")
	End Sub
	
	Private Sub EventSImage_GotFocus()
		Call PfRaiseEvent("GotFocus")
	End Sub
	
	Private Sub EventSImage_KeyDown(ByRef KeyCode As Short, ByRef Shift As Short)
		Call PfRaiseEvent("KeyDown", KeyCode, Shift)
	End Sub
	
	Private Sub EventSImage_KeyPress(ByRef KeyAscii As Short)
		Call PfRaiseEvent("KeyPress", KeyAscii)
	End Sub
	
	Private Sub EventSImage_KeyUp(ByRef KeyCode As Short, ByRef Shift As Short)
		Call PfRaiseEvent("KeyUp", KeyCode, Shift)
	End Sub
	
	Private Sub EventSImage_LostFocus()
		Call PfRaiseEvent("LostFocus")
	End Sub
	
	Private Sub EventSImage_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles EventSImage.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call PfRaiseEvent("MouseDown", Button, Shift, x, y)
		Call obj_MouseDown(EventSImage, Button, Shift, x, y)
	End Sub
	
	Private Sub EventSImage_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles EventSImage.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call PfRaiseEvent("MouseMove", Button, Shift, x, y)
		Call obj_MouseMove(EventSImage, Button, Shift, x, y)
	End Sub
	
	Private Sub EventSImage_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles EventSImage.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call PfRaiseEvent("MouseUp", Button, Shift, x, y)
	End Sub
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Dynamic Image Event
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	Private Sub EventDImage_Change()
		Call PfRaiseEvent("Change")
	End Sub
	
	Private Sub EventDImage_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EventDImage.Click
		Call PfRaiseEvent("Click")
		Call obj_Click(EventDImage, 3)
	End Sub
	
	Private Sub EventDImage_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EventDImage.DoubleClick
		Call PfRaiseEvent("DblClick")
	End Sub
	
	Private Sub EventDImage_GotFocus()
		Call PfRaiseEvent("GotFocus")
	End Sub
	
	Private Sub EventDImage_KeyDown(ByRef KeyCode As Short, ByRef Shift As Short)
		Call PfRaiseEvent("KeyDown", KeyCode, Shift)
	End Sub
	
	Private Sub EventDImage_KeyPress(ByRef KeyAscii As Short)
		Call PfRaiseEvent("KeyPress", KeyAscii)
	End Sub
	
	Private Sub EventDImage_KeyUp(ByRef KeyCode As Short, ByRef Shift As Short)
		Call PfRaiseEvent("KeyUp", KeyCode, Shift)
	End Sub
	
	Private Sub EventDImage_LostFocus()
		Call PfRaiseEvent("LostFocus")
	End Sub
	
	Private Sub EventDImage_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles EventDImage.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call PfRaiseEvent("MouseDown", Button, Shift, x, y)
		Call obj_MouseDown(EventDImage, Button, Shift, x, y)
	End Sub
	
	Private Sub EventDImage_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles EventDImage.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call PfRaiseEvent("MouseMove", Button, Shift, x, y)
		Call obj_MouseMove(EventDImage, Button, Shift, x, y)
	End Sub
	
	Private Sub EventDImage_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles EventDImage.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call PfRaiseEvent("MouseUp", Button, Shift, x, y)
	End Sub
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Line Image Event
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	Private Sub EventLImage_Change()
		Call PfRaiseEvent("Change")
	End Sub
	
	Private Sub EventLImage_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EventLImage.Click
		Call PfRaiseEvent("Click")
		Call obj_Click(EventLImage, 5)
	End Sub
	
	Private Sub EventLImage_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EventLImage.DoubleClick
		Call PfRaiseEvent("DblClick")
	End Sub
	
	Private Sub EventLImage_GotFocus()
		Call PfRaiseEvent("GotFocus")
	End Sub
	
	Private Sub EventLImage_KeyDown(ByRef KeyCode As Short, ByRef Shift As Short)
		Call PfRaiseEvent("KeyDown", KeyCode, Shift)
	End Sub
	
	Private Sub EventLImage_KeyPress(ByRef KeyAscii As Short)
		Call PfRaiseEvent("KeyPress", KeyAscii)
	End Sub
	
	Private Sub EventLImage_KeyUp(ByRef KeyCode As Short, ByRef Shift As Short)
		Call PfRaiseEvent("KeyUp", KeyCode, Shift)
	End Sub
	
	Private Sub EventLImage_LostFocus()
		Call PfRaiseEvent("LostFocus")
	End Sub
	
	Private Sub EventLImage_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles EventLImage.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call PfRaiseEvent("MouseDown", Button, Shift, x, y)
		Call obj_MouseDown(EventLImage, Button, Shift, x, y)
	End Sub
	
	Private Sub EventLImage_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles EventLImage.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call PfRaiseEvent("MouseMove", Button, Shift, x, y)
		Call obj_MouseMove(EventLImage, Button, Shift, x, y)
	End Sub
	
	Private Sub EventLImage_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles EventLImage.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call PfRaiseEvent("MouseUp", Button, Shift, x, y)
	End Sub
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Barcode Image Event
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	Private Sub EventBImage_Change()
		Call PfRaiseEvent("Change")
	End Sub
	
	Private Sub EventBImage_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EventBImage.Click
		Call PfRaiseEvent("Click")
		Call obj_Click(EventBImage, 4)
	End Sub
	
	Private Sub EventBImage_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EventBImage.DoubleClick
		Call PfRaiseEvent("DblClick")
	End Sub
	
	Private Sub EventBImage_GotFocus()
		Call PfRaiseEvent("GotFocus")
	End Sub
	
	Private Sub EventBImage_KeyDown(ByRef KeyCode As Short, ByRef Shift As Short)
		Call PfRaiseEvent("KeyDown", KeyCode, Shift)
	End Sub
	
	Private Sub EventBImage_KeyPress(ByRef KeyAscii As Short)
		Call PfRaiseEvent("KeyPress", KeyAscii)
	End Sub
	
	Private Sub EventBImage_KeyUp(ByRef KeyCode As Short, ByRef Shift As Short)
		Call PfRaiseEvent("KeyUp", KeyCode, Shift)
	End Sub
	
	Private Sub EventBImage_LostFocus()
		Call PfRaiseEvent("LostFocus")
	End Sub
	
	Private Sub EventBImage_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles EventBImage.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call PfRaiseEvent("MouseDown", Button, Shift, x, y)
		Call obj_MouseDown(EventBImage, Button, Shift, x, y)
	End Sub
	
	Private Sub EventBImage_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles EventBImage.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call PfRaiseEvent("MouseMove", Button, Shift, x, y)
		Call obj_MouseMove(EventBImage, Button, Shift, x, y)
	End Sub
	
	Private Sub EventBImage_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles EventBImage.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Call PfRaiseEvent("MouseUp", Button, Shift, x, y)
	End Sub
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Line Event
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	Private Sub EventLine_Change()
		Call PfRaiseEvent("Change")
	End Sub
	
	Private Sub EventLine_Click()
		Call PfRaiseEvent("Click")
		Call obj_Click(EventLine, 0)
	End Sub
	
	Private Sub EventLine_DblClick()
		Call PfRaiseEvent("DblClick")
	End Sub
	
	Private Sub EventLine_GotFocus()
		Call PfRaiseEvent("GotFocus")
	End Sub
	
	Private Sub EventLine_KeyDown(ByRef KeyCode As Short, ByRef Shift As Short)
		Call PfRaiseEvent("KeyDown", KeyCode, Shift)
	End Sub
	
	Private Sub EventLine_KeyPress(ByRef KeyAscii As Short)
		Call PfRaiseEvent("KeyPress", KeyAscii)
	End Sub
	
	Private Sub EventLine_KeyUp(ByRef KeyCode As Short, ByRef Shift As Short)
		Call PfRaiseEvent("KeyUp", KeyCode, Shift)
	End Sub
	
	Private Sub EventLine_LostFocus()
		Call PfRaiseEvent("LostFocus")
	End Sub
	
	Private Sub EventLine_MouseDown(ByRef Button As Short, ByRef Shift As Short, ByRef x As Single, ByRef y As Single)
		Call PfRaiseEvent("MouseDown", Button, Shift, x, y)
		Call obj_MouseDown(EventLine, Button, Shift, x, y)
	End Sub
	
	Private Sub EventLine_MouseMove(ByRef Button As Short, ByRef Shift As Short, ByRef x As Single, ByRef y As Single)
		Call PfRaiseEvent("MouseMove", Button, Shift, x, y)
		Call obj_MouseMove(EventLine, Button, Shift, x, y)
	End Sub
	
	Private Sub EventLine_MouseUp(ByRef Button As Short, ByRef Shift As Short, ByRef x As Single, ByRef y As Single)
		Call PfRaiseEvent("MouseUp", Button, Shift, x, y)
	End Sub
	
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Barcode Event
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	Private Sub EventBarcode_Change()
		Call PfRaiseEvent("Change")
	End Sub
	
	Private Sub EventBarcode_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EventBarcode.ClickEvent
		Call PfRaiseEvent("Click")
		Call obj_Click(EventBarcode, 4)
	End Sub
	
	Private Sub EventBarcode_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EventBarcode.DblClick
		Call PfRaiseEvent("DblClick")
	End Sub
	
	Private Sub EventBarcode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EventBarcode.Enter
		Call PfRaiseEvent("GotFocus")
	End Sub
	
	Private Sub EventBarcode_KeyDown(ByRef KeyCode As Short, ByRef Shift As Short)
		Call PfRaiseEvent("KeyDown", KeyCode, Shift)
	End Sub
	
	Private Sub EventBarcode_KeyPress(ByRef KeyAscii As Short)
		Call PfRaiseEvent("KeyPress", KeyAscii)
	End Sub
	
	Private Sub EventBarcode_KeyUp(ByRef KeyCode As Short, ByRef Shift As Short)
		Call PfRaiseEvent("KeyUp", KeyCode, Shift)
	End Sub
	
	Private Sub EventBarcode_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EventBarcode.Leave
		Call PfRaiseEvent("LostFocus")
	End Sub
	
	Private Sub EventBarcode_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxBarcodLib._DBarcodEvents_MouseDownEvent) Handles EventBarcode.MouseDownEvent
		Call PfRaiseEvent("MouseDown", eventArgs.Button, eventArgs.Shift, eventArgs.x, eventArgs.y)
		Call obj_MouseDown(EventBarcode, eventArgs.Button, eventArgs.Shift, eventArgs.x, eventArgs.y)
	End Sub
	
	Private Sub EventBarcode_MouseMoveEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxBarcodLib._DBarcodEvents_MouseMoveEvent) Handles EventBarcode.MouseMoveEvent
		Call PfRaiseEvent("MouseMove", eventArgs.Button, eventArgs.Shift, eventArgs.x, eventArgs.y)
		Call obj_MouseMove(EventBarcode, eventArgs.Button, eventArgs.Shift, eventArgs.x, eventArgs.y)
	End Sub
	
	Private Sub EventBarcode_MouseUpEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxBarcodLib._DBarcodEvents_MouseUpEvent) Handles EventBarcode.MouseUpEvent
		Call PfRaiseEvent("MouseUp", eventArgs.Button, eventArgs.Shift, eventArgs.x, eventArgs.y)
	End Sub
End Class