VERSION 5.00
Begin VB.UserControl Barcode 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3195
   LockControls    =   -1  'True
   ScaleHeight     =   588
   ScaleMode       =   0  '사용자
   ScaleWidth      =   3195
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '평면
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   651.568
      ScaleMode       =   0  '사용자
      ScaleWidth      =   3135
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "Barcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim nhDC As Long
Dim cColor As Long
Dim n
Dim ConstDC(37), ConstLADC(10), ConstRDC(10), LHGDC, CGPDC, RHGDC As String
Dim ConstBDC(10) As String
Dim WonStr As String
Dim k
Dim Insu As String

'기본 속성 값:
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
Const m_def_Caption = "880"
Const m_def_CodeStyle = "EAN-8"
'속성 변수:
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
Dim m_Caption As String
Dim m_CodeStyle As String
'이벤트 선언:
Event Click()
Attribute Click.VB_Description = "개체에서 마우스 단추를 눌렀다가 놓을 때 발생합니다."
Event DblClick()
Attribute DblClick.VB_Description = "마우스 단추를 개체에서 누르고 놓은 후 다시 누르고 놓으면 발생합니다."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "개체에 포커스가 있을 때 키를 누르면 발생합니다."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "ANSI키를 누르고 놓았을 경우 발생합니다."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "개체에 포커스가 있을 때 키를 놓으면 발생합니다."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Attribute MouseDown.VB_Description = "개체에 포커스가 있을 때 마우스 단추를 누르면 발생합니다."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Attribute MouseMove.VB_Description = "마우스를 움직일 경우 발생합니다."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
Attribute MouseUp.VB_Description = "개체에 포커스가 있을 때 마우스 단추를 놓으면 발생합니다."
Public Property Get BackColor() As Long
Attribute BackColor.VB_Description = "개체의 텍스트나 그래픽을 표시하기 위해 사용되는 배경색을 반환하거나 설정합니다."
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property
Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "개체에서 텍스트나 그래픽을 표시하는 전경색을 반환하거나 설정합니다."
    ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "사용자가 만든 이벤트에 대해 개체가 응답할 수 있는지의 여부를 결정하는 값을 반환하거나 설정합니다."
    Enabled = m_Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property
Public Property Get Font() As Font
Attribute Font.VB_Description = "Font 개체를 반환합니다."
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Label이나 Shape의 배경이 투명 또는 불투명한지의 여부를 나타냅니다."
    BackStyle = m_BackStyle
End Property
Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "개체 테두리 유형을 반환하거나 설정합니다."
    BorderStyle = m_BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property
Public Sub Refresh()
Attribute Refresh.VB_Description = "개체를 완전히 다시 그리게 합니다."
     
End Sub
Public Property Get Caption() As String
    Caption = m_Caption
End Property
Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    Select Case CodeStyle
       Case "Code39"
         Barcode_Set   '3of9 바코드 폰트 정의 및 출력함수
       Case "EAN-8"
         EAN_Set       'EAN-8 바코드 폰트 정의 및 출력함수
       Case "EAN-13"
         EAN13_Set
       Case "UPC-A"
         UPCA_Set
    End Select
    
'    If CodeStyle = "Code 3 of 9" Then
'       Barcode_Set   '3of9 바코드 폰트 정의 및 출력함수
'    ElseIf CodeStyle = "EAN-8" Then
'       EAN_Set       'EAN-8 바코드 폰트 정의 및 출력함수
'    End If
End Property
Public Property Get CodeStyle() As String
    CodeStyle = m_CodeStyle
End Property
Public Property Let CodeStyle(ByVal New_CodeStyle As String)
    m_CodeStyle = New_CodeStyle
    PropertyChanged "CodeStyle"
End Property
'프로퍼티값 초기화 기술
'사용자 정의 컨트롤 속성을 초기화함.
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
    m_Caption = m_def_Caption
    m_CodeStyle = m_def_CodeStyle
End Sub
'저장 장치에서 속성 값을 불러온다.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_CodeStyle = PropBag.ReadProperty("CodeStyle", m_def_CodeStyle)
End Sub
'저장 장치에 속성 값을 쓴다.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("CodeStyle", m_CodeStyle, m_def_CodeStyle)
End Sub
'3of9 바코드 폰트 정의 및 출력함수
Private Sub Barcode_Set()
Dim ConvAsc As Integer

   ConstDC(10) = "1001011011010"
   ConstDC(0) = "1010011011010"
   ConstDC(1) = "1101001010110"
   ConstDC(2) = "1011001010110"
   ConstDC(3) = "1101100101010"
   ConstDC(4) = "1010011010110"
   ConstDC(5) = "1101001101010"
   ConstDC(6) = "1011001101010"
   ConstDC(7) = "1010010110110"
   ConstDC(8) = "1101001011010"
   ConstDC(9) = "1011001011010"
   ConstDC(11) = "1101010010110"
   ConstDC(12) = "1011010010110"
   ConstDC(13) = "1101101001010"
   ConstDC(14) = "1010110010110"
   ConstDC(15) = "1101011001010"
   ConstDC(16) = "1011011001010"
   ConstDC(17) = "1010100110110"
   ConstDC(18) = "1101010011010"
   ConstDC(19) = "1011010011010"
   ConstDC(20) = "1010110011010"
   ConstDC(21) = "1101010100110"
   ConstDC(22) = "1011010100110"
   ConstDC(23) = "1101101010010"
   ConstDC(24) = "1010110100110"
   ConstDC(25) = "1101011010010"
   ConstDC(26) = "1011011010010"
   ConstDC(27) = "1011011010010"
   ConstDC(28) = "1101010110010"
   ConstDC(29) = "1011010110010"
   ConstDC(30) = "1010110110010"
   ConstDC(31) = "1100101010110"
   ConstDC(32) = "1001101010110"
   ConstDC(33) = "1100110101010"
   ConstDC(34) = "1001011010110"
   ConstDC(35) = "1100101101010"
   ConstDC(36) = "1001101101010"
   
  If WonStr <> Caption Then
    WonStr = Caption
    WonStr = UCase(WonStr)
    Insu = ""
    Insu = ConstDC(10)   '3of9바코드 시작값 삽입
    For k = 0 To Len(WonStr) - 1
       '바코그가 영문자 일경우
       If Asc(Mid(WonStr, k + 1, 1)) > 64 And Asc(Mid(WonStr, k + 1, 1)) < 91 Then
          ConvAsc = Asc(Mid(WonStr, k + 1, 1))
          Insu = Insu & ConstDC(ConvAsc - 54)
       Else  '바코드가 숫자일경우
          Insu = Insu & ConstDC(Val(Mid(WonStr, k + 1, 1)))
       End If
    Next
    Insu = Insu & ConstDC(10)   '3of9바코드 종료값 삽입
    
    Bar_Display Insu, Len(WonStr)
  End If
End Sub
'바코드를 컨트롤에 출력한다
Function Bar_Display(Str As String, StrLen As Long)
Dim DanChr As String
Dim i
Dim XPosition As Long
Dim g
    
    XPosition = 2
    Picture1.Cls
    
       For i = 0 To Len(Str) - 1
          DanChr = Mid(Str, i + 1, 1)
          If StrLen > 6 And CodeStyle = "Code39" Then  '바코드가 6자리 이상일경우 1줄씩 인쇄
             If DanChr = 1 Then
                Picture1.Line (XPosition * 15, 0)-(XPosition * 15, 820)
             Else
                '
             End If
                XPosition = XPosition + 1
          Else   '바코드가 6자리 이하 일경우 2줄씩 인쇄
             If DanChr = 1 Then
                Picture1.Line (XPosition * 15, 0)-(XPosition * 15, 820)
                Picture1.Line ((XPosition + 1) * 15, 0)-((XPosition + 1) * 15, 820)
             Else
                '
                '
             End If
                XPosition = XPosition + 2
          End If
       Next
      
     ' BarImage = Picture1.Image
      
End Function
'EAN-8 바코드 폰트 정의 및 출력함수
Private Sub EAN_Set()
Dim ConvAsc As Integer

   LHGDC = "101"
   CGPDC = "01010"
   RHGDC = "101"
   
   ConstLADC(0) = "0001101"
   ConstLADC(1) = "0011001"
   ConstLADC(2) = "0010011"
   ConstLADC(3) = "0111101"
   ConstLADC(4) = "0100011"
   ConstLADC(5) = "0110001"
   ConstLADC(6) = "0101111"
   ConstLADC(7) = "0111011"
   ConstLADC(8) = "0110111"
   ConstLADC(9) = "0001011"
   
   ConstRDC(0) = "1110010"
   ConstRDC(1) = "1100110"
   ConstRDC(2) = "1101100"
   ConstRDC(3) = "1000010"
   ConstRDC(4) = "1011100"
   ConstRDC(5) = "1001110"
   ConstRDC(6) = "1010000"
   ConstRDC(7) = "1000100"
   ConstRDC(8) = "1001000"
   ConstRDC(9) = "1110100"
   
      
  If WonStr <> Caption Then
    WonStr = Caption
    'WonStr = UCase(WonStr)
    Insu = ""
    Insu = LHGDC   'EAN-8 바코드 시작값 삽입
    For k = 0 To Len(WonStr) - 1
       If k = 4 Then Insu = Insu & CGPDC
       If k < 4 Then  'Left Side 숫자일 경우
          Insu = Insu & ConstLADC(Val(Mid(WonStr, k + 1, 1)))
       Else           'Right Side 숫자일 경우
          Insu = Insu & ConstRDC(Val(Mid(WonStr, k + 1, 1)))
       End If
    Next
    Insu = Insu & RHGDC   'EAN-8 바코드 종료값 삽입
    
    Bar_Display Insu, Len(WonStr)
  End If
End Sub
'EAN-13 바코드 폰트 정의 및 출력함수
Private Sub EAN13_Set()
Dim ConvAsc As Integer
Dim g As Integer
Dim FF As String

   LHGDC = "101"
   CGPDC = "01010"
   RHGDC = "101"
   
   ConstLADC(0) = "0001101"
   ConstLADC(1) = "0011001"
   ConstLADC(2) = "0010011"
   ConstLADC(3) = "0111101"
   ConstLADC(4) = "0100011"
   ConstLADC(5) = "0110001"
   ConstLADC(6) = "0101111"
   ConstLADC(7) = "0111011"
   ConstLADC(8) = "0110111"
   ConstLADC(9) = "0001011"
   
   ConstRDC(0) = "1110010"
   ConstRDC(1) = "1100110"
   ConstRDC(2) = "1101100"
   ConstRDC(3) = "1000010"
   ConstRDC(4) = "1011100"
   ConstRDC(5) = "1001110"
   ConstRDC(6) = "1010000"
   ConstRDC(7) = "1000100"
   ConstRDC(8) = "1001000"
   ConstRDC(9) = "1110100"
   
   ConstBDC(0) = "0100111"
   ConstBDC(1) = "0110011"
   ConstBDC(2) = "0011011"
   ConstBDC(3) = "0100001"
   ConstBDC(4) = "0011101"
   ConstBDC(5) = "0111001"
   ConstBDC(6) = "0000101"
   ConstBDC(7) = "0010001"
   ConstBDC(8) = "0001001"
   ConstBDC(9) = "0010111"
      
  If WonStr <> Caption Then
    WonStr = Caption
    'WonStr = UCase(WonStr)
    Insu = ""
    Insu = LHGDC   'EAN-13 바코드 시작값 삽입
    
    For g = 0 To Len(WonStr) - 1
       Select Case g
         Case 0
            FF = Left(WonStr, 1) 'First Flag 발췌
         Case 1
            Insu = Insu & ConstLADC(Val(Mid(WonStr, g + 1, 1)))
         Case 2
            Select Case FF
              Case "0"
                 Insu = Insu & ConstLADC(Val(Mid(WonStr, g + 1, 1)))
              Case "1"
                 Insu = Insu & ConstLADC(Val(Mid(WonStr, g + 1, 1)))
              Case "2"
                 Insu = Insu & ConstLADC(Val(Mid(WonStr, g + 1, 1)))
              Case "3"
                 Insu = Insu & ConstLADC(Val(Mid(WonStr, g + 1, 1)))
              Case "4"
                 Insu = Insu & ConstBDC(Val(Mid(WonStr, g + 1, 1)))
              Case "5"
                 Insu = Insu & ConstBDC(Val(Mid(WonStr, g + 1, 1)))
              Case "6"
                 Insu = Insu & ConstBDC(Val(Mid(WonStr, g + 1, 1)))
              Case "7"
                 Insu = Insu & ConstBDC(Val(Mid(WonStr, g + 1, 1)))
              Case "8"
                 Insu = Insu & ConstBDC(Val(Mid(WonStr, g + 1, 1)))
              Case "9"
                 Insu = Insu & ConstBDC(Val(Mid(WonStr, g + 1, 1)))
            End Select
         Case 3
            Select Case FF
              Case "0"
                 Insu = Insu & ConstLADC(Val(Mid(WonStr, g + 1, 1)))
              Case "1"
                 Insu = Insu & ConstBDC(Val(Mid(WonStr, g + 1, 1)))
              Case "2"
                 Insu = Insu & ConstBDC(Val(Mid(WonStr, g + 1, 1)))
              Case "3"
                 Insu = Insu & ConstBDC(Val(Mid(WonStr, g + 1, 1)))
              Case "4"
                 Insu = Insu & ConstLADC(Val(Mid(WonStr, g + 1, 1)))
              Case "5"
                 Insu = Insu & ConstBDC(Val(Mid(WonStr, g + 1, 1)))
              Case "6"
                 Insu = Insu & ConstBDC(Val(Mid(WonStr, g + 1, 1)))
              Case "7"
                 Insu = Insu & ConstLADC(Val(Mid(WonStr, g + 1, 1)))
              Case "8"
                 Insu = Insu & ConstLADC(Val(Mid(WonStr, g + 1, 1)))
              Case "9"
                 Insu = Insu & ConstBDC(Val(Mid(WonStr, g + 1, 1)))
            End Select
         Case 4
            Select Case FF
              Case "0"
                 Insu = Insu & ConstLADC(Val(Mid(WonStr, g + 1, 1)))
              Case "1"
                 Insu = Insu & ConstLADC(Val(Mid(WonStr, g + 1, 1)))
              Case "2"
                 Insu = Insu & ConstBDC(Val(Mid(WonStr, g + 1, 1)))
              Case "3"
                 Insu = Insu & ConstBDC(Val(Mid(WonStr, g + 1, 1)))
              Case "4"
                 Insu = Insu & ConstLADC(Val(Mid(WonStr, g + 1, 1)))
              Case "5"
                 Insu = Insu & ConstLADC(Val(Mid(WonStr, g + 1, 1)))
              Case "6"
                 Insu = Insu & ConstBDC(Val(Mid(WonStr, g + 1, 1)))
              Case "7"
                 Insu = Insu & ConstBDC(Val(Mid(WonStr, g + 1, 1)))
              Case "8"
                 Insu = Insu & ConstBDC(Val(Mid(WonStr, g + 1, 1)))
              Case "9"
                 Insu = Insu & ConstLADC(Val(Mid(WonStr, g + 1, 1)))
            End Select
         Case 5
            Select Case FF
              Case "0"
                 Insu = Insu & ConstLADC(Val(Mid(WonStr, g + 1, 1)))
              Case "1"
                 Insu = Insu & ConstBDC(Val(Mid(WonStr, g + 1, 1)))
              Case "2"
                 Insu = Insu & ConstLADC(Val(Mid(WonStr, g + 1, 1)))
              Case "3"
                 Insu = Insu & ConstBDC(Val(Mid(WonStr, g + 1, 1)))
              Case "4"
                 Insu = Insu & ConstBDC(Val(Mid(WonStr, g + 1, 1)))
              Case "5"
                 Insu = Insu & ConstLADC(Val(Mid(WonStr, g + 1, 1)))
              Case "6"
                 Insu = Insu & ConstLADC(Val(Mid(WonStr, g + 1, 1)))
              Case "7"
                 Insu = Insu & ConstLADC(Val(Mid(WonStr, g + 1, 1)))
              Case "8"
                 Insu = Insu & ConstBDC(Val(Mid(WonStr, g + 1, 1)))
              Case "9"
                 Insu = Insu & ConstBDC(Val(Mid(WonStr, g + 1, 1)))
            End Select
         Case 6
            Select Case FF
              Case "0"
                 Insu = Insu & ConstLADC(Val(Mid(WonStr, g + 1, 1)))
              Case "1"
                 Insu = Insu & ConstBDC(Val(Mid(WonStr, g + 1, 1)))
              Case "2"
                 Insu = Insu & ConstBDC(Val(Mid(WonStr, g + 1, 1)))
              Case "3"
                 Insu = Insu & ConstLADC(Val(Mid(WonStr, g + 1, 1)))
              Case "4"
                 Insu = Insu & ConstBDC(Val(Mid(WonStr, g + 1, 1)))
              Case "5"
                 Insu = Insu & ConstBDC(Val(Mid(WonStr, g + 1, 1)))
              Case "6"
                 Insu = Insu & ConstLADC(Val(Mid(WonStr, g + 1, 1)))
              Case "7"
                 Insu = Insu & ConstBDC(Val(Mid(WonStr, g + 1, 1)))
              Case "8"
                 Insu = Insu & ConstLADC(Val(Mid(WonStr, g + 1, 1)))
              Case "9"
                 Insu = Insu & ConstLADC(Val(Mid(WonStr, g + 1, 1)))
            End Select
         Case 7
            Insu = Insu & CGPDC
            Insu = Insu & ConstRDC(Val(Mid(WonStr, g + 1, 1)))
         Case 8
            Insu = Insu & ConstRDC(Val(Mid(WonStr, g + 1, 1)))
         Case 9
            Insu = Insu & ConstRDC(Val(Mid(WonStr, g + 1, 1)))
         Case 10
            Insu = Insu & ConstRDC(Val(Mid(WonStr, g + 1, 1)))
         Case 11
            Insu = Insu & ConstRDC(Val(Mid(WonStr, g + 1, 1)))
         Case 12
            Insu = Insu & ConstRDC(Val(Mid(WonStr, g + 1, 1)))
       End Select
    Next
    
    Insu = Insu & RHGDC   'EAN-13 바코드 종료값 삽입
    
    Bar_Display Insu, Len(WonStr)
  End If
End Sub
'UPC-A 바코드 폰트 정의 및 출력함수
Private Sub UPCA_Set()
Dim ConvAsc As Integer
Dim y As Integer

   LHGDC = "101"
   CGPDC = "01010"
   RHGDC = "101"
   
   ConstLADC(0) = "0001101"
   ConstLADC(1) = "0011001"
   ConstLADC(2) = "0010011"
   ConstLADC(3) = "0111101"
   ConstLADC(4) = "0100011"
   ConstLADC(5) = "0110001"
   ConstLADC(6) = "0101111"
   ConstLADC(7) = "0111011"
   ConstLADC(8) = "0110111"
   ConstLADC(9) = "0001011"
   
   ConstRDC(0) = "1110010"
   ConstRDC(1) = "1100110"
   ConstRDC(2) = "1101100"
   ConstRDC(3) = "1000010"
   ConstRDC(4) = "1011100"
   ConstRDC(5) = "1001110"
   ConstRDC(6) = "1010000"
   ConstRDC(7) = "1000100"
   ConstRDC(8) = "1001000"
   ConstRDC(9) = "1110100"
   
      
  If WonStr <> Caption Then
    WonStr = Caption
    'WonStr = UCase(WonStr)
    Insu = ""
    Insu = LHGDC   'EAN-8 바코드 시작값 삽입
    For y = 0 To Len(WonStr) - 1
       If y = 6 Then Insu = Insu & CGPDC
       If y < 6 Then  'Left Side 숫자일 경우
          Insu = Insu & ConstLADC(Val(Mid(WonStr, y + 1, 1)))
       Else           'Right Side 숫자일 경우
          Insu = Insu & ConstRDC(Val(Mid(WonStr, y + 1, 1)))
       End If
    Next
    Insu = Insu & RHGDC   'EAN-8 바코드 종료값 삽입
    
    Bar_Display Insu, Len(WonStr)
  End If
End Sub
'경고! 다음 행들을 없애거나 바꾸지마!..죽음뿐
Public Property Get Image() As Picture
Attribute Image.VB_Description = "Microsoft Windows에 의해 제공된 핸들을 비트맵에 지속적으로 반환합니다."
    Set Image = Picture1.Image
End Property

