VERSION 5.00
Begin VB.UserControl MDButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1380
   DefaultCancel   =   -1  'True
   MaskColor       =   &H00000000&
   ScaleHeight     =   26
   ScaleMode       =   3  '픽셀
   ScaleWidth      =   92
   ToolboxBitmap   =   "MD_Button.ctx":0000
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   270
      Index           =   2
      Left            =   585
      Picture         =   "MD_Button.ctx":0312
      Top             =   75
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   315
      Picture         =   "MD_Button.ctx":0744
      Top             =   75
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   270
      Index           =   0
      Left            =   45
      Picture         =   "MD_Button.ctx":0B76
      Top             =   75
      Visible         =   0   'False
      Width           =   270
   End
End
Attribute VB_Name = "MDButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' **********************************************************************************************
' Date Created      : 2003.06.19
' Initial Author    : 윤태현
' E-Mail            : thyoon@mdsaver.net
' Purpose           : Button 컨트롤
' Description       :
' Dependencies      :
' **********************************************************************************************
Option Explicit

Enum ButtonStyle '버튼의 Draw Style
   NormalButton = 0     '보통 이미지
   ClickButton = 1      '마우스 Down 이미지
   DisabledButton = 2   'Disable된 이미지
End Enum

Const m_def_Caption = "&Ok" '기본 Caption
Dim m_Caption As String
Dim m_Style As ButtonStyle
'이벤트 선언:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "개체에서 마우스 단추를 눌렀다가 놓을 때 발생합니다."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "마우스 단추를 개체에서 누르고 놓은 후 다시 누르고 놓으면 발생합니다."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "개체에 포커스가 있을 때 키를 누르면 발생합니다."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "ANSI키를 누르고 놓았을 경우 발생합니다."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "개체에 포커스가 있을 때 키를 놓으면 발생합니다."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "개체에 포커스가 있을 때 마우스 단추를 누르면 발생합니다."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "마우스를 움직일 경우 발생합니다."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "개체에 포커스가 있을 때 마우스 단추를 놓으면 발생합니다."

'=========================================================
'이 아래부터는 컨트롤 Mapping임
'=========================================================

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "사용자가 만든 이벤트에 대해 개체가 응답할 수 있는지의 여부를 결정하는 값을 반환하거나 설정합니다."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    DrawButton

End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Font 개체를 반환합니다."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    RaiseEvent Click
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    DrawButton ClickButton
    RaiseEvent MouseDown(Button, Shift, X, y)
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    DrawButton NormalButton
    RaiseEvent MouseUp(Button, Shift, X, y)
End Sub

Public Property Get Caption() As String
   Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
   m_Caption = New_Caption
   DrawButton
   PropertyChanged "Caption"
End Property

'사용자 정의 컨트롤에 대한 속성을 초기화합니다.
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_Caption = m_def_Caption
End Sub

Private Sub UserControl_Resize()
    DrawButton
End Sub

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "사용자 정의 마우스 아이콘을 설정합니다."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "개체에 마우스 포인터가 있을 경우 표시되는 마우스 포인터의 형식을 반환하거나 설정합니다."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property


'=========================================================
'이 아래부터는 속성값 처리
'=========================================================

'저장소에서 속성값을 로드합니다.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    DrawButton
End Sub


'속성값을 저장소에 기록합니다.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
End Sub

'=========================================================
'이 아래부터는 내부함수 처리
'=========================================================

'********************************************************************************************
'Purpose       : 버튼 표시
'--------------------------------------------------------------------------------------------
'Developer     Date        Comments
'--------------------------------------------------------------------------------------------
'윤태현        2003-06-19  최초작성자
'********************************************************************************************
Private Sub DrawButton(Optional nStyle As ButtonStyle)
Dim nW As Long, nH As Long
Dim nX As Long, nY As Long
Dim nAT As Long
Dim cTemp As String
Dim cLeft As String

Dim brx, bry, bw, bh As Integer

    '비활성화시는 Click, Normal은 무시된다.
    If UserControl.Enabled = False Then
        m_Style = DisabledButton
    Else
        m_Style = nStyle
    End If
    With UserControl
        UserControl.ScaleMode = 3 'pixels로 구정
        'Short cuts
        brx = UserControl.ScaleWidth - 6
        bry = UserControl.ScaleHeight - 6
        bw = UserControl.ScaleWidth - 12
        bh = UserControl.ScaleHeight - 12
        '버튼 그리기
        With Image1(m_Style)
            UserControl.Cls
            UserControl.PaintPicture .Picture, 0, 0, 6, 6, 0, 0, 6, 6           '상좌 Corner
            UserControl.PaintPicture .Picture, brx, 0, 6, 6, 12, 0, 6, 6        '상우
            UserControl.PaintPicture .Picture, brx, bry, 6, 6, 12, 12, 6, 6     '우하
            UserControl.PaintPicture .Picture, 0, bry, 6, 6, 0, 12, 6, 6        '좌하
            
            UserControl.PaintPicture .Picture, 6, 0, bw, 6, 6, 0, 6, 6          'Top
            UserControl.PaintPicture .Picture, brx, 6, 6, bh, 12, 6, 6, 6       'Right
            UserControl.PaintPicture .Picture, 0, 6, 6, bh, 0, 6, 6, 6          'Left
            UserControl.PaintPicture .Picture, 6, bry, bw, 6, 6, 12, 6, 6       'bottom
            UserControl.PaintPicture .Picture, 6, 6, bw, bh, 6, 6, 6, 6         'Background
        End With
        
        '캡션에 &를 표시하기 위한 &&를 우선 Tab으로 변경
        cTemp = Replace(m_Caption, "&&", Chr(9))
        'HotKey 지정
        nAT = InStr(1, cTemp, "&")
        If nAT <> 0 Then
            .AccessKeys = Mid(cTemp, nAT + 1, 1)
        End If
        '& 삭제
        cTemp = Replace(cTemp, "&", "")
        '&& -> Tab -> & 로 표시
        cTemp = Replace(cTemp, Chr(9), "&")
        'Caption 위치 구함
        nW = .TextWidth(cTemp)
        nH = .TextHeight(cTemp)
        nX = (UserControl.ScaleWidth - nW) / 2
        nY = (UserControl.ScaleHeight - nH) / 2
        .CurrentX = nX
        .CurrentY = nY
        
        'Caption 표시
        UserControl.Print cTemp
        
        'Hotkey _표시
        If nAT <> 0 Then
            cLeft = Left(cTemp, nAT)
            .CurrentX = nX + .TextWidth(cLeft) - .TextWidth(Right(cLeft, 1))
            .CurrentY = nY + nH
            UserControl.Line (.CurrentX, .CurrentY)-(.CurrentX + .TextWidth(Right(cLeft, 1)), .CurrentY)
        End If
        .Refresh
    End With
End Sub

