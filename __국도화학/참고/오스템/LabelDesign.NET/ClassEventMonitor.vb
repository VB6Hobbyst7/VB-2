Option Strict Off
Option Explicit On
Friend Class ClassEventMonitor
	'===============================================================================
	'  ���α׷� : ���� ���� Object���� �߻��� �̺�Ʈ ������ ���� Ŭ���� ���
	'  �� �� �� : ClassEventMonitor.cls
	'  �� �� �� : 2007.03.30
	'  �� �� �� : ������(182cm@korea.com)
	'  Ȩ������ : http://www.EnjoyDev.com
	'  ��    �� :
	'
	'  �����̷�
	'===============================================================================
	'  Flag    ��������    ������   ��������
	'-------------------------------------------------------------------------------
	'  [CYJ#0] 2007.03.30
	'  [OSW]   2011.09.21  ������   ��,�̹���,����,���ڵ� �߰�
	'===============================================================================
	
	
	'-- �̺�Ʈ
	Public Event EventRaised(ByRef EventObject As ClassEventObject, ByVal EventName As String)
	
	'-- Enum :: ���� ���� ��Ʈ�� ID
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
	'UPGRADE_NOTE: Class_Initialize��(��) Class_Initialize_Renamed(��)�� ���׷��̵�Ǿ����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'UPGRADE_NOTE: Class_Terminate��(��) Class_Terminate_Renamed(��)�� ���׷��̵�Ǿ����ϴ�. �ڼ��� ������ ������ �����Ͻʽÿ�. 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	
	'===============================================================================
	' �� �� �� : RaiseUserEvent()
	' ��    �� : �̺�Ʈ ����
	' �� �� �� :
	' �� �� �� :
	' �� �� �� : 2007.03.30
	' �� �� �� : ������(182cm@korea.com)
	'===============================================================================
	Public Sub RaiseUserEvent(ByRef EventObject As ClassEventObject, ByVal EventName As String)
		
		RaiseEvent EventRaised(EventObject, EventName)
		
	End Sub
End Class