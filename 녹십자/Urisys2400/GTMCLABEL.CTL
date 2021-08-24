VERSION 5.00
Begin VB.UserControl gtmcLabel 
   ClientHeight    =   1290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2145
   ControlContainer=   -1  'True
   ScaleHeight     =   1290
   ScaleWidth      =   2145
   Begin VB.Shape shpBorder 
      BorderStyle     =   0  '투명
      Height          =   345
      Left            =   135
      Top             =   45
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   720
      TabIndex        =   0
      Top             =   135
      Width           =   540
   End
   Begin VB.Line linBottom 
      BorderColor     =   &H80000006&
      BorderWidth     =   2
      X1              =   105
      X2              =   540
      Y1              =   825
      Y2              =   825
   End
   Begin VB.Line linRight 
      BorderWidth     =   2
      X1              =   180
      X2              =   555
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line linLeft 
      X1              =   165
      X2              =   495
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Line linTop 
      X1              =   165
      X2              =   480
      Y1              =   540
      Y2              =   540
   End
End
Attribute VB_Name = "gtmcLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'기본 속성 값:
Const cRatio = 3
Const cGap = 3

Const cDark = &H80000003
Const cBack = &H8000000F
Const cBright = &H80000005
Const cBlack = &H80000008
 
Const m_def_Border = 0
Const m_def_BevelStyle = 0
Const m_def_TextPosition = 5

'const m_def_Black

Public Enum TextPosition
    TopLeft = 1
    TopCenter = 2
    TopRight = 3
    MiddleLeft = 4
    MiddleCenter = 5
    MiddleRight = 6
    BottomLeft = 7
    BottomCenter = 8
    BottomRight = 9
End Enum

Public Enum BDStyle
    None = 0
    singleborder = 1
End Enum

Public Enum BevelStyle
    Flat = 0
    Convex = 1
    Concave = 2
End Enum

'속성 변수:
Dim m_Border As BDStyle
Dim m_BevelSize As Integer
Dim m_BevelStyle As BevelStyle
Dim m_TextPosition As TextPosition

'이벤트 선언:
Event Click() 'MappingInfo=Label1,Label1,-1,Click
Attribute Click.VB_Description = "개체에서 마우스 단추를 눌렀다가 놓을 때 발생합니다."
Event DblClick() 'MappingInfo=Label1,Label1,-1,DblClick
Attribute DblClick.VB_Description = "마우스 단추를 개체에서 누르고 놓은 후 다시 누르고 놓으면 발생합니다."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "개체가 포커스를 가질 때 키를 누르면 발생합니다."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "ANSI 키를 누르고 놓았을 경우 발생합니다."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "개체가 포커스를 가질 때 키를 놓으면 발생합니다."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Label1,Label1,-1,MouseDown
Attribute MouseDown.VB_Description = "개체에 포커스가 있을 때 키를 누르면 발생합니다."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Label1,Label1,-1,MouseMove
Attribute MouseMove.VB_Description = "마우스를 움직일 경우 발생합니다."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Label1,Label1,-1,MouseUp
Attribute MouseUp.VB_Description = "개체에 포커스를 맞추고 있는 동안 마우스 단추를 놓았을 경우 발생합니다."


'경고! 다음에 설명하는 행들을 없애거나 바꾸지 마십시오!
'MappingInfo=Label1,Label1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "개체의 텍스트나 그래픽을 표시하기위해 사용되는 배경색을 반환하거나 설정합니다."
    BackColor = Label1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Label1.BackColor() = New_BackColor
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'경고! 다음에 설명하는 행들을 없애거나 바꾸지 마십시오!
'MappingInfo=Label1,Label1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "개체에서 텍스트나 그래픽을 표시하는 전경색을 반환하거나 설정합니다."
    ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Label1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'경고! 다음에 설명하는 행들을 없애거나 바꾸지 마십시오!
'MappingInfo=Label1,Label1,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "개체가 사용자가 만든 이벤트에 대해 응답할 수 있는지의 여부를 결정하는 값을 반환하거나 설정합니다."
    Enabled = Label1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Label1.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'경고! 다음에 설명하는 행들을 없애거나 바꾸지 마십시오!
'MappingInfo=Label1,Label1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Font 개체를 반환합니다."
Attribute Font.VB_UserMemId = -512
    Set Font = Label1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get BevelStyle() As BevelStyle
Attribute BevelStyle.VB_Description = "개체 테두리 유형을 반환하거나 설정합니다."
    BevelStyle = m_BevelStyle
End Property

Public Property Let BevelStyle(ByVal New_BevelStyle As BevelStyle)
    m_BevelStyle = New_BevelStyle
    SetBevelStyle
    PropertyChanged "BevelStyle"
End Property

'경고! 다음에 설명하는 행들을 없애거나 바꾸지 마십시오!
'MappingInfo=Label1,Label1,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "개체를 완전히 다시 그리게 합니다."
    Label1.Refresh
End Sub

Private Sub Label1_Click()
    RaiseEvent Click
End Sub

Private Sub Label1_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Property Get Border() As BDStyle
    Border = m_Border
End Property

Public Property Let Border(ByVal New_Border As BDStyle)
    m_Border = New_Border
    SetBorder
    PropertyChanged "Border"
End Property

Public Property Get TextPosition() As TextPosition
Attribute TextPosition.VB_Description = "레이블 내의 텍스트 위치 지정"
    TextPosition = m_TextPosition
End Property

Public Property Let TextPosition(ByVal New_TextPosition As TextPosition)
    m_TextPosition = New_TextPosition
    SetTextPosition     ' 위치 설정
    PropertyChanged "TextPosition"
End Property

'사용자 정의 컨트롤 속성을 초기화합니다.
Private Sub UserControl_InitProperties()

    m_Border = m_def_Border
    m_BevelStyle = m_def_BevelStyle
    m_TextPosition = m_def_TextPosition
    
End Sub

'저장 장치에서 속성 값을 불러옵니다.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Label1.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Label1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Label1.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Border = PropBag.ReadProperty("Border", m_Border)
    m_BevelStyle = PropBag.ReadProperty("BevelStyle", m_def_BevelStyle)
    m_TextPosition = PropBag.ReadProperty("TextPosition", m_def_TextPosition)
        
    Label1.Caption = PropBag.ReadProperty("Caption", "DrLabel")
    
End Sub

Private Sub UserControl_Resize()
    
    'Label1.Top = 0: Label1.Left = 0
    'Label1.Width = usercontrol.Width
    'Label1.Height = usercontrol.Height
    
    SetBorderSize
    'SetBevelSize
    SetBevelStyle
    SetTextPosition
        
End Sub

Private Sub SetBorder()
    
    Select Case m_Border
        Case None:          shpBorder.BorderStyle = 0 'Visible = False
        Case singleborder:  shpBorder.BorderStyle = 1 'Visible = True
    End Select
    
End Sub

Private Sub SetTextPosition()
    
    Label1.AutoSize = True
    
    Select Case m_TextPosition
        
        Case TopLeft: Label1.Move linLeft.X1 + cRatio + 100, linTop.Y1 + cRatio + 100
        Case TopCenter: Label1.Move (UserControl.Width - Label1.Width) / 2, linTop.Y1 + cRatio + 100
        Case TopRight: Label1.Move linTop.X2 - Label1.Width - cRatio - 100, linTop.Y1 + cRatio + 100
        
        Case MiddleLeft: Label1.Move linLeft.X1 + cRatio + 100, (UserControl.Height - Label1.Height) / 2
        Case MiddleCenter: Label1.Move (UserControl.Width - Label1.Width) / 2, (UserControl.Height - Label1.Height) / 2
        Case MiddleRight: Label1.Move linTop.X2 - Label1.Width - cRatio - 100, (UserControl.Height - Label1.Height) / 2
        
        Case BottomLeft: Label1.Move linLeft.X1 + cRatio + 100, linBottom.Y1 - Label1.Height - cRatio - 100
        Case BottomCenter: Label1.Move (UserControl.Width - Label1.Width) / 2, linBottom.Y1 - Label1.Height - cRatio - 100
        Case BottomRight: Label1.Move linTop.X2 - Label1.Width - cRatio - 100, linBottom.Y1 - Label1.Height - cRatio - 100
        
    End Select
    
    UserControl.AutoRedraw = True
    
End Sub

Private Sub SetBevelStyle()
Dim I As Integer, sVal As Single

    sVal = 3

    Label1.AutoSize = True

    Select Case m_BevelStyle

        Case Flat
                MoveLine linTop, cRatio, cRatio, UserControl.Width - cRatio * sVal, cRatio, cBack
                MoveLine linLeft, linTop.X1, linTop.Y1, linTop.X1, UserControl.Height - cRatio * sVal, cBack
                MoveLine linRight, linTop.X2 - cGap, linTop.Y2, linTop.X2 - cGap, UserControl.Height - cRatio * sVal, cBack
                MoveLine linBottom, linLeft.X2, linLeft.Y2 - cGap, linRight.X2, linRight.Y2 - cGap, cBack
                linTop.Visible = False
                linLeft.Visible = False
                linRight.Visible = False
                linBottom.Visible = False
        Case Convex
                MoveLine linTop, cRatio, cRatio, UserControl.Width - cRatio * sVal, cRatio, cDark
                MoveLine linLeft, linTop.X1, linTop.Y1, linTop.X1, UserControl.Height - cRatio * sVal, cDark
                MoveLine linRight, linTop.X2 - cGap, linTop.Y2, linTop.X2 - cGap, UserControl.Height - cRatio * sVal, cBright
                MoveLine linBottom, linLeft.X2, linLeft.Y2 - cGap, linRight.X2, linRight.Y2 - cGap, cBright
                linTop.Visible = True
                linLeft.Visible = True
                linRight.Visible = True
                linBottom.Visible = True
        Case Concave
                MoveLine linTop, cRatio, cRatio, UserControl.Width - cRatio * sVal, cRatio, cBright
                MoveLine linLeft, linTop.X1, linTop.Y1, linTop.X1, UserControl.Height - cRatio * sVal, cBright
                MoveLine linRight, linTop.X2 - cGap, linTop.Y2, linTop.X2 - cGap, UserControl.Height - cRatio * sVal, cDark
                MoveLine linBottom, linLeft.X2, linLeft.Y2 - cGap, linRight.X2, linRight.Y2 - cGap, cDark
                linTop.Visible = True
                linLeft.Visible = True
                linRight.Visible = True
                linBottom.Visible = True
    End Select
    
    linTop.ZOrder 0
    linLeft.ZOrder 0

End Sub


Private Sub SetBorderSize()
    shpBorder.Top = 0: shpBorder.Left = 0
    shpBorder.Width = UserControl.Width
    shpBorder.Height = UserControl.Height
End Sub

Private Sub MoveLine(ByVal Obj As Line, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal Color As Long)

    Obj.X1 = X1
    Obj.Y1 = Y1
    Obj.X2 = X2
    Obj.Y2 = Y2
    Obj.BorderColor = Color

End Sub

'저장 장치에 속성 값을 씁니다.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", Label1.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", Label1.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", Label1.Enabled, True)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("Border", m_Border, m_def_Border)
    Call PropBag.WriteProperty("BevelStyle", m_BevelStyle, m_def_BevelStyle)
    Call PropBag.WriteProperty("TextPosition", m_TextPosition, m_def_TextPosition)
    
    Call PropBag.WriteProperty("Caption", Label1.Caption, "gtmcLabel")
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=Label1,Label1,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "개체의 제목 표시줄이나 개체 아이콘 밑에 표시되는 텍스트를 반환하거나 설정합니다."
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption() = New_Caption
    SetTextPosition
    PropertyChanged "Caption"
End Property

