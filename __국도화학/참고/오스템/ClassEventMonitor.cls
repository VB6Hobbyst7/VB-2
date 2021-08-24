VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassEventMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
Option Explicit


'-- 이벤트
Public Event EventRaised(ByRef EventObject As ClassEventObject, ByVal EventName As String)

'-- Enum :: 동적 생성 컨트롤 ID
Public Enum EventObjectID
    EventObjectCommandButton            ' CommandButton
'    EventObjectTextBox                  ' TextBox
    EventObjectTextBox                  ' TextBox
    EventObjectSLabel                   ' Static_Label
    EventObjectDLabel                   ' Dynamic_Label
    EventObjectBLabel                   ' Barcode_Label
    EventObjectSImage                   ' Static_Image
    EventObjectDImage                   ' Dynamic_Image
    EventObjectLImage                   ' Line_Image
    EventObjectBImage                   ' Barcode_Image
    EventObjectBarcode                  ' Barcode
    EventObjectLine                     ' Line
    EventObjectRLabel                   ' Dynamic_Label - RFID
End Enum

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Class_Initialize()

End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Class_Terminate()

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


