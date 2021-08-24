Option Strict Off
Option Explicit On
Friend Class ClassEventMonitor
	'===============================================================================
	'  프로그램 : 동적 생성 Object에서 발생한 이벤트 전달을 위한 클래스 모듈
	'  파 일 명 : ClassEventMonitor.cls
	'  작 성 일 : 2007.03.30
	'  작 성 자 : 제용재(182cm@korea.com)
	'  홈페이지 : http://www.EnjoyDev.com
	'  설    명 :
	'
	'  수정이력
	'===============================================================================
	'  Flag    수정일자    수정자   수정내용
	'-------------------------------------------------------------------------------
	'  [CYJ#0] 2007.03.30
	'  [OSW]   2011.09.21  오세원   라벨,이미지,라인,바코드 추가
	'===============================================================================
	
	
	'-- 이벤트
	Public Event EventRaised(ByRef EventObject As ClassEventObject, ByVal EventName As String)
	
	'-- Enum :: 동적 생성 컨트롤 ID
	Public Enum EventObjectID
		EventObjectCommandButton ' CommandButton
		EventObjectTextBox ' TextBox
		EventObjectSLabel ' Static_Label
		EventObjectDLabel ' Dynamic_Label
		EventObjectBLabel ' Barcode_Label
		EventObjectSImage ' Static_Image
		EventObjectDImage ' Dynamic_Image
		EventObjectLImage ' Line_Image
		EventObjectBImage ' Barcode_Image
		EventObjectBarcode ' Barcode
		EventObjectLine ' Line
	End Enum
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'UPGRADE_NOTE: Class_Initialize이(가) Class_Initialize_Renamed(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'UPGRADE_NOTE: Class_Terminate이(가) Class_Terminate_Renamed(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	
	'===============================================================================
	' 함 수 명 : RaiseUserEvent()
	' 설    명 : 이벤트 전달
	' 입 력 값 :
	' 결 과 값 :
	' 작 성 일 : 2007.03.30
	' 작 성 자 : 제용재(182cm@korea.com)
	'===============================================================================
	Public Sub RaiseUserEvent(ByRef EventObject As ClassEventObject, ByVal EventName As String)
		
		RaiseEvent EventRaised(EventObject, EventName)
		
	End Sub
End Class