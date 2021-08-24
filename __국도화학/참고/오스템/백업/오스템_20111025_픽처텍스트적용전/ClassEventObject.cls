VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassEventObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
Option Explicit

Private m_ClsEventMonitor               As ClassEventMonitor        ' 이벤트 전달을 위한 클래스
Private m_FrmOwner                      As frmLabelDesign           ' 부모 폼 정보
Private m_IntEventObjectId              As EventObjectID            ' CommandButton, TextBox, Label, Image ...
Private m_VarParam()                    As Variant                  ' 이벤트 발생에서의 파라미터 정보

' 이벤트 처리를 위한 Object
Private WithEvents EventCommandButton   As VB.CommandButton
Attribute EventCommandButton.VB_VarHelpID = -1
Private WithEvents EventTextBox         As VB.textbox
Attribute EventTextBox.VB_VarHelpID = -1
Private WithEvents EventSLabel          As VB.Label
Attribute EventSLabel.VB_VarHelpID = -1
Private WithEvents EventDLabel          As VB.Label
Attribute EventDLabel.VB_VarHelpID = -1
Private WithEvents EventBLabel          As VB.Label
Attribute EventBLabel.VB_VarHelpID = -1
Private WithEvents EventSImage          As VB.Image
Attribute EventSImage.VB_VarHelpID = -1
Private WithEvents EventDImage          As VB.Image
Attribute EventDImage.VB_VarHelpID = -1
Private WithEvents EventLImage          As VB.Image
Attribute EventLImage.VB_VarHelpID = -1
Private WithEvents EventBImage          As VB.Image
Attribute EventBImage.VB_VarHelpID = -1
Private WithEvents EventLine            As VB.Line
Attribute EventLine.VB_VarHelpID = -1
Private WithEvents EventBarcode         As BarcodLib.Barcod
Attribute EventBarcode.VB_VarHelpID = -1

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' 클래스 생성자
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Class_Initialize()

    ' 클래스 초기화
    Set EventCommandButton = Nothing
    Set EventTextBox = Nothing
    Set EventSLabel = Nothing
    Set EventDLabel = Nothing
    Set EventBLabel = Nothing
    Set EventSImage = Nothing
    Set EventDImage = Nothing
    Set EventLImage = Nothing
    Set EventBImage = Nothing
    Set EventLine = Nothing
    Set EventBarcode = Nothing

End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' 클래스 소멸자
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Class_Terminate()

    On Error Resume Next
    
    ' 동적 생성 컨트롤 제거
    m_FrmOwner.Controls.Remove Me.EventObject.Name

    Set EventCommandButton = Nothing
    Set EventTextBox = Nothing
    Set EventSLabel = Nothing
    Set EventDLabel = Nothing
    Set EventBLabel = Nothing
    Set EventSImage = Nothing
    Set EventDImage = Nothing
    Set EventLImage = Nothing
    Set EventBImage = Nothing
    Set EventLine = Nothing
    Set EventBarcode = Nothing

End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' EventMonitor Property
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Property Set EventMonitor(ByRef ClsEventMonitor As ClassEventMonitor)
    Set m_ClsEventMonitor = ClsEventMonitor
End Property

Public Property Get EventMonitor()
    Set EventMonitor = m_ClsEventMonitor
End Property

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Owner Property
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Property Set Owner(ByRef frm As Form)
    Set m_FrmOwner = frm
End Property

Public Property Get Owner()
    Set Owner = m_FrmOwner
End Property

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' EventObject Property
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Property Get EventObject()

    Select Case m_IntEventObjectId
        Case EventObjectCommandButton
            Set EventObject = EventCommandButton
            
        Case EventObjectTextBox
            Set EventObject = EventTextBox

        Case EventObjectSLabel
            Set EventObject = EventSLabel

        Case EventObjectDLabel
            Set EventObject = EventDLabel

        Case EventObjectBLabel
            Set EventObject = EventBLabel

        Case EventObjectSImage
            Set EventObject = EventSImage
        
        Case EventObjectDImage
            Set EventObject = EventDImage

        Case EventObjectLImage
            Set EventObject = EventLImage

        Case EventObjectBImage
            Set EventObject = EventBImage

        Case EventObjectLine
            Set EventObject = EventLine

        Case EventObjectBarcode
            Set EventObject = EventBarcode

    End Select
    
End Property


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Param Property
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Property Get Param(ByVal IntIndex As Integer) As Variant

    On Error Resume Next
    Param = m_VarParam(IntIndex)

End Property


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Param Property
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Property Get Params() As Variant

    Params = m_VarParam

End Property


'===============================================================================
' 함 수 명 : CreateObject()
' 설    명 : 동적 컨트롤 생성
' 입 력 값 :
' 결 과 값 :
'===============================================================================
Public Function CreateObject(ByRef FrmOwner As Form, ByRef ClsEventMonitor As ClassEventMonitor, ByVal IntEventObjectId As EventObjectID, ByVal StrObjectName As String) As Object
    Dim obj             As Object
'    Dim objB            As BarcodLib.Barcod

    On Error Resume Next

    Set m_FrmOwner = FrmOwner
    Set m_ClsEventMonitor = ClsEventMonitor
        
    m_IntEventObjectId = IntEventObjectId
    
    Select Case IntEventObjectId

        ' CommandButton
        Case EventObjectCommandButton
            Set obj = m_FrmOwner.Controls.Add("VB.CommandButton", StrObjectName)
            Set EventCommandButton = obj

        ' TextBox
        Case EventObjectTextBox
            Set obj = m_FrmOwner.Controls.Add("VB.TextBox", StrObjectName)
            Set EventTextBox = obj
        
        ' Static_Label
        Case EventObjectSLabel
            Set obj = m_FrmOwner.Controls.Add("VB.Label", StrObjectName)
            Set EventSLabel = obj
        
        ' Dynamic_Label
        Case EventObjectDLabel
            Set obj = m_FrmOwner.Controls.Add("VB.Label", StrObjectName)
            Set EventDLabel = obj
        
        ' Barcode_Label
        Case EventObjectBLabel
            Set obj = m_FrmOwner.Controls.Add("VB.Label", StrObjectName)
            Set EventBLabel = obj
        
        ' Static_Image
        Case EventObjectSImage
            Set obj = m_FrmOwner.Controls.Add("VB.Image", StrObjectName)
            Set EventSImage = obj
        
        ' Dynamic_Image
        Case EventObjectDImage
            Set obj = m_FrmOwner.Controls.Add("VB.Image", StrObjectName)
            Set EventDImage = obj
        
        ' Line_Image
        Case EventObjectLImage
            Set obj = m_FrmOwner.Controls.Add("VB.Image", StrObjectName)
            Set EventLImage = obj
        
        ' Barcode_Image
        Case EventObjectBImage
            Set obj = m_FrmOwner.Controls.Add("VB.Image", StrObjectName)
            Set EventBImage = obj
        
        ' Line
        Case EventObjectLine
            Set obj = m_FrmOwner.Controls.Add("VB.Line", StrObjectName)
            Set EventLine = obj
        
        ' Barcode
        Case EventObjectBarcode
            Set obj = m_FrmOwner.Controls.Add("Mabry.BarCod", StrObjectName)
            Set EventBarcode = obj
            
        Case Else
            MsgBox "EventObjectId Error!!", vbCritical
            Exit Function

    End Select

    Set CreateObject = obj
    
End Function

'===============================================================================
' 함 수 명 : PfRaiseEvent()
' 설    명 : 이벤트 발생 시킴
' 입 력 값 :
' 결 과 값 :
' 작 성 일 : 2007.03.30
' 작 성 자 : 제용재(182cm@korea.com)
'===============================================================================
Private Sub PfRaiseEvent(ByVal StrEventName As String, ParamArray Params() As Variant)

    ' 파라미터 저장
    m_VarParam = Params

    ' CommandButton의 Click 이벤트 전달
    Call m_ClsEventMonitor.RaiseUserEvent(Me, StrEventName)

End Sub



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' CommandButton Event
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub EventCommandButton_Click()
    Call PfRaiseEvent("Click")
End Sub

Private Sub EventCommandButton_GotFocus()
    Call PfRaiseEvent("GotFocus")
End Sub

Private Sub EventCommandButton_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PfRaiseEvent("KeyDown", KeyCode, Shift)
End Sub

Private Sub EventCommandButton_KeyPress(KeyAscii As Integer)
    Call PfRaiseEvent("KeyPress", KeyAscii)
End Sub

Private Sub EventCommandButton_KeyUp(KeyCode As Integer, Shift As Integer)
    Call PfRaiseEvent("KeyUp", KeyCode, Shift)
End Sub

Private Sub EventCommandButton_LostFocus()
    Call PfRaiseEvent("LostFocus")
End Sub

Private Sub EventCommandButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseDown", Button, Shift, x, y)
End Sub

Private Sub EventCommandButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseMove", Button, Shift, x, y)
End Sub

Private Sub EventCommandButton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseUp", Button, Shift, x, y)
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' TextBox Event
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub EventTextBox_Change()
    Call PfRaiseEvent("Change")
End Sub

Private Sub EventTextBox_Click()
    Call PfRaiseEvent("Click")
End Sub

Private Sub EventTextBox_DblClick()
    Call PfRaiseEvent("DblClick")
End Sub

Private Sub EventTextBox_GotFocus()
    Call PfRaiseEvent("GotFocus")
End Sub

Private Sub EventTextBox_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PfRaiseEvent("KeyDown", KeyCode, Shift)
End Sub

Private Sub EventTextBox_KeyPress(KeyAscii As Integer)
    Call PfRaiseEvent("KeyPress", KeyAscii)
End Sub

Private Sub EventTextBox_KeyUp(KeyCode As Integer, Shift As Integer)
    Call PfRaiseEvent("KeyUp", KeyCode, Shift)
End Sub

Private Sub EventTextBox_LostFocus()
    Call PfRaiseEvent("LostFocus")
End Sub

Private Sub EventTextBox_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseDown", Button, Shift, x, y)
End Sub

Private Sub EventTextBox_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseMove", Button, Shift, x, y)
End Sub

Private Sub EventTextBox_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseUp", Button, Shift, x, y)
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Static Label Event
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub EventSLabel_Change()
    Call PfRaiseEvent("Change")
End Sub

Private Sub EventSLabel_Click()
    Call PfRaiseEvent("Click")
    Call obj_Click(EventSLabel, 0)
End Sub

Private Sub EventSLabel_DblClick()
    Call PfRaiseEvent("DblClick")
End Sub

Private Sub EventSLabel_GotFocus()
    Call PfRaiseEvent("GotFocus")
End Sub

Private Sub EventSLabel_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PfRaiseEvent("KeyDown", KeyCode, Shift)
End Sub

Private Sub EventSLabel_KeyPress(KeyAscii As Integer)
    Call PfRaiseEvent("KeyPress", KeyAscii)
End Sub

Private Sub EventSLabel_KeyUp(KeyCode As Integer, Shift As Integer)
    Call PfRaiseEvent("KeyUp", KeyCode, Shift)
End Sub

Private Sub EventSLabel_LostFocus()
    Call PfRaiseEvent("LostFocus")
End Sub

Private Sub EventSLabel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseDown", Button, Shift, x, y)
    Call obj_MouseDown(EventSLabel, Button, Shift, x, y)
End Sub

Private Sub EventSLabel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseMove", Button, Shift, x, y)
    Call obj_MouseMove(EventSLabel, Button, Shift, x, y)
End Sub

Private Sub EventSLabel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseUp", Button, Shift, x, y)
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Dynamic Label Event
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub EventDLabel_Change()
    Call PfRaiseEvent("Change")
End Sub

Private Sub EventDLabel_Click()
    Call PfRaiseEvent("Click")
    Call obj_Click(EventDLabel, 1)
End Sub

Private Sub EventDLabel_DblClick()
    Call PfRaiseEvent("DblClick")
End Sub

Private Sub EventDLabel_GotFocus()
    Call PfRaiseEvent("GotFocus")
End Sub

Private Sub EventDLabel_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PfRaiseEvent("KeyDown", KeyCode, Shift)
End Sub

Private Sub EventDLabel_KeyPress(KeyAscii As Integer)
    Call PfRaiseEvent("KeyPress", KeyAscii)
End Sub

Private Sub EventDLabel_KeyUp(KeyCode As Integer, Shift As Integer)
    Call PfRaiseEvent("KeyUp", KeyCode, Shift)
End Sub

Private Sub EventDLabel_LostFocus()
    Call PfRaiseEvent("LostFocus")
End Sub

Private Sub EventDLabel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseDown", Button, Shift, x, y)
    Call obj_MouseDown(EventDLabel, Button, Shift, x, y)
End Sub

Private Sub EventDLabel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseMove", Button, Shift, x, y)
    Call obj_MouseMove(EventDLabel, Button, Shift, x, y)
End Sub

Private Sub EventDLabel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseUp", Button, Shift, x, y)
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Barcode Label Event
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub EventBLabel_Change()
    Call PfRaiseEvent("Change")
End Sub

Private Sub EventBLabel_Click()
    Call PfRaiseEvent("Click")
    Call obj_Click(EventBLabel, 4)
End Sub

Private Sub EventBLabel_DblClick()
    Call PfRaiseEvent("DblClick")
End Sub

Private Sub EventBLabel_GotFocus()
    Call PfRaiseEvent("GotFocus")
End Sub

Private Sub EventBLabel_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PfRaiseEvent("KeyDown", KeyCode, Shift)
End Sub

Private Sub EventBLabel_KeyPress(KeyAscii As Integer)
    Call PfRaiseEvent("KeyPress", KeyAscii)
End Sub

Private Sub EventBLabel_KeyUp(KeyCode As Integer, Shift As Integer)
    Call PfRaiseEvent("KeyUp", KeyCode, Shift)
End Sub

Private Sub EventBLabel_LostFocus()
    Call PfRaiseEvent("LostFocus")
End Sub

Private Sub EventBLabel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent(EventBLabel, "MouseDown", Button, Shift, x, y)
    Call obj_MouseDown(EventBLabel, Button, Shift, x, y)
End Sub

Private Sub EventBLabel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseMove", Button, Shift, x, y)
    Call obj_MouseMove(EventBLabel, Button, Shift, x, y)
End Sub

Private Sub EventBLabel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseUp", Button, Shift, x, y)
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Static Image Event
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub EventSImage_Change()
    Call PfRaiseEvent("Change")
End Sub

Private Sub EventSImage_Click()
    Call PfRaiseEvent("Click")
    Call obj_Click(EventSImage, 2)
End Sub

Private Sub EventSImage_DblClick()
    Call PfRaiseEvent("DblClick")
End Sub

Private Sub EventSImage_GotFocus()
    Call PfRaiseEvent("GotFocus")
End Sub

Private Sub EventSImage_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PfRaiseEvent("KeyDown", KeyCode, Shift)
End Sub

Private Sub EventSImage_KeyPress(KeyAscii As Integer)
    Call PfRaiseEvent("KeyPress", KeyAscii)
End Sub

Private Sub EventSImage_KeyUp(KeyCode As Integer, Shift As Integer)
    Call PfRaiseEvent("KeyUp", KeyCode, Shift)
End Sub

Private Sub EventSImage_LostFocus()
    Call PfRaiseEvent("LostFocus")
End Sub

Private Sub EventSImage_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseDown", Button, Shift, x, y)
    Call obj_MouseDown(EventSImage, Button, Shift, x, y)
End Sub

Private Sub EventSImage_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseMove", Button, Shift, x, y)
    Call obj_MouseMove(EventSImage, Button, Shift, x, y)
End Sub

Private Sub EventSImage_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseUp", Button, Shift, x, y)
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Dynamic Image Event
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub EventDImage_Change()
    Call PfRaiseEvent("Change")
End Sub

Private Sub EventDImage_Click()
    Call PfRaiseEvent("Click")
    Call obj_Click(EventDImage, 3)
End Sub

Private Sub EventDImage_DblClick()
    Call PfRaiseEvent("DblClick")
End Sub

Private Sub EventDImage_GotFocus()
    Call PfRaiseEvent("GotFocus")
End Sub

Private Sub EventDImage_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PfRaiseEvent("KeyDown", KeyCode, Shift)
End Sub

Private Sub EventDImage_KeyPress(KeyAscii As Integer)
    Call PfRaiseEvent("KeyPress", KeyAscii)
End Sub

Private Sub EventDImage_KeyUp(KeyCode As Integer, Shift As Integer)
    Call PfRaiseEvent("KeyUp", KeyCode, Shift)
End Sub

Private Sub EventDImage_LostFocus()
    Call PfRaiseEvent("LostFocus")
End Sub

Private Sub EventDImage_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseDown", Button, Shift, x, y)
    Call obj_MouseDown(EventDImage, Button, Shift, x, y)
End Sub

Private Sub EventDImage_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseMove", Button, Shift, x, y)
    Call obj_MouseMove(EventDImage, Button, Shift, x, y)
End Sub

Private Sub EventDImage_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseUp", Button, Shift, x, y)
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Line Image Event
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub EventLImage_Change()
    Call PfRaiseEvent("Change")
End Sub

Private Sub EventLImage_Click()
    Call PfRaiseEvent("Click")
    Call obj_Click(EventLImage, 5)
End Sub

Private Sub EventLImage_DblClick()
    Call PfRaiseEvent("DblClick")
End Sub

Private Sub EventLImage_GotFocus()
    Call PfRaiseEvent("GotFocus")
End Sub

Private Sub EventLImage_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PfRaiseEvent("KeyDown", KeyCode, Shift)
End Sub

Private Sub EventLImage_KeyPress(KeyAscii As Integer)
    Call PfRaiseEvent("KeyPress", KeyAscii)
End Sub

Private Sub EventLImage_KeyUp(KeyCode As Integer, Shift As Integer)
    Call PfRaiseEvent("KeyUp", KeyCode, Shift)
End Sub

Private Sub EventLImage_LostFocus()
    Call PfRaiseEvent("LostFocus")
End Sub

Private Sub EventLImage_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseDown", Button, Shift, x, y)
    Call obj_MouseDown(EventLImage, Button, Shift, x, y)
End Sub

Private Sub EventLImage_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseMove", Button, Shift, x, y)
    Call obj_MouseMove(EventLImage, Button, Shift, x, y)
End Sub

Private Sub EventLImage_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseUp", Button, Shift, x, y)
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Barcode Image Event
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub EventBImage_Change()
    Call PfRaiseEvent("Change")
End Sub

Private Sub EventBImage_Click()
    Call PfRaiseEvent("Click")
    Call obj_Click(EventBImage, 4)
End Sub

Private Sub EventBImage_DblClick()
    Call PfRaiseEvent("DblClick")
End Sub

Private Sub EventBImage_GotFocus()
    Call PfRaiseEvent("GotFocus")
End Sub

Private Sub EventBImage_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PfRaiseEvent("KeyDown", KeyCode, Shift)
End Sub

Private Sub EventBImage_KeyPress(KeyAscii As Integer)
    Call PfRaiseEvent("KeyPress", KeyAscii)
End Sub

Private Sub EventBImage_KeyUp(KeyCode As Integer, Shift As Integer)
    Call PfRaiseEvent("KeyUp", KeyCode, Shift)
End Sub

Private Sub EventBImage_LostFocus()
    Call PfRaiseEvent("LostFocus")
End Sub

Private Sub EventBImage_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseDown", Button, Shift, x, y)
    Call obj_MouseDown(EventBImage, Button, Shift, x, y)
End Sub

Private Sub EventBImage_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseMove", Button, Shift, x, y)
    Call obj_MouseMove(EventBImage, Button, Shift, x, y)
End Sub

Private Sub EventBImage_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
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

Private Sub EventLine_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PfRaiseEvent("KeyDown", KeyCode, Shift)
End Sub

Private Sub EventLine_KeyPress(KeyAscii As Integer)
    Call PfRaiseEvent("KeyPress", KeyAscii)
End Sub

Private Sub EventLine_KeyUp(KeyCode As Integer, Shift As Integer)
    Call PfRaiseEvent("KeyUp", KeyCode, Shift)
End Sub

Private Sub EventLine_LostFocus()
    Call PfRaiseEvent("LostFocus")
End Sub

Private Sub EventLine_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseDown", Button, Shift, x, y)
    Call obj_MouseDown(EventLine, Button, Shift, x, y)
End Sub

Private Sub EventLine_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseMove", Button, Shift, x, y)
    Call obj_MouseMove(EventLine, Button, Shift, x, y)
End Sub

Private Sub EventLine_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseUp", Button, Shift, x, y)
End Sub


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Barcode Event
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub EventBarcode_Change()
    Call PfRaiseEvent("Change")
End Sub

Private Sub EventBarcode_Click()
    Call PfRaiseEvent("Click")
    Call obj_Click(EventBarcode, 4)
End Sub

Private Sub EventBarcode_DblClick()
    Call PfRaiseEvent("DblClick")
End Sub

Private Sub EventBarcode_GotFocus()
    Call PfRaiseEvent("GotFocus")
End Sub

Private Sub EventBarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PfRaiseEvent("KeyDown", KeyCode, Shift)
End Sub

Private Sub EventBarcode_KeyPress(KeyAscii As Integer)
    Call PfRaiseEvent("KeyPress", KeyAscii)
End Sub

Private Sub EventBarcode_KeyUp(KeyCode As Integer, Shift As Integer)
    Call PfRaiseEvent("KeyUp", KeyCode, Shift)
End Sub

Private Sub EventBarcode_LostFocus()
    Call PfRaiseEvent("LostFocus")
End Sub

Private Sub EventBarcode_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseDown", Button, Shift, x, y)
    Call obj_MouseDown(EventBarcode, Button, Shift, x, y)
End Sub

Private Sub EventBarcode_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseMove", Button, Shift, x, y)
    Call obj_MouseMove(EventBarcode, Button, Shift, x, y)
End Sub

Private Sub EventBarcode_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PfRaiseEvent("MouseUp", Button, Shift, x, y)
End Sub


