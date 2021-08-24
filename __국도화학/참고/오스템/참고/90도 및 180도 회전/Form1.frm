VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5085
   ClientLeft      =   9360
   ClientTop       =   3240
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5085
   ScaleWidth      =   7290
   Begin VB.CommandButton Command2 
      Caption         =   "리셋"
      Height          =   390
      Left            =   150
      TabIndex        =   4
      Top             =   1650
      Width           =   1965
   End
   Begin VB.CommandButton Command1 
      Caption         =   "180도 회전"
      Height          =   390
      Index           =   2
      Left            =   150
      TabIndex        =   3
      Top             =   1125
      Width           =   1965
   End
   Begin VB.CommandButton Command1 
      Caption         =   "-90도 회전"
      Height          =   390
      Index           =   1
      Left            =   150
      TabIndex        =   2
      Top             =   675
      Width           =   1965
   End
   Begin VB.CommandButton Command1 
      Caption         =   "90도 회전"
      Height          =   390
      Index           =   0
      Left            =   150
      TabIndex        =   1
      Top             =   225
      Width           =   1965
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '평면
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   2325
      ScaleHeight     =   279
      ScaleMode       =   3  '픽셀
      ScaleWidth      =   275
      TabIndex        =   0
      Top             =   135
      Width           =   4155
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
    On Error Resume Next
    Dim w1 As Long, h1 As Long '크기
    Dim w2 As Long, h2 As Long '크기
    
    Dim x As Long, y As Long '변수
    Dim xx As Long, yy As Long '변수
    Dim r As Single, g As Single, b As Single
    
    '원본 이미지 설정
    w1 = Picture1.ScaleWidth
    h1 = Picture1.ScaleHeight
    
    '비트맵 배열을 준비한다.
    Call SetBitmapArray_Input(w1, h1)
    
    '이미지를 비트맵 배열로 복사한다.
    Call GetDIBits(Picture1.hdc, Picture1.Image, 0, h1, BITMAP_INPUT(0, 0, 0), BITMAP_INFO, DIB_RGB_COLORS)
    
    '비트맵 배열을 준비한다.
    If Index = 0 Or Index = 1 Then '90도, -90도 회전
        '출력 이미지 설정
        w2 = Picture1.ScaleHeight
        h2 = Picture1.ScaleWidth
        
        Call SetBitmapArray_Output(w2, h2)
    Else '180도 회전
        '출력 이미지 설정
        w2 = Picture1.ScaleWidth
        h2 = Picture1.ScaleHeight
        
        Call SetBitmapArray_Output(w2, h2)
    End If
    
    '배열 처리
    For y = 0 To h2 - 1
        For x = 0 To w2 - 1
        
            Select Case Index
            Case 1 '시계방향 90도 회전
                xx = (h2 - 1) - y
                yy = x
            Case 0 '반시계방향 90도 회전
                xx = y
                yy = (w2 - 1) - x
            Case 2 '180도 회전
                xx = (w2 - 1) - x
                yy = (h2 - 1) - y
            End Select
            
            'INPUT, OUTPUT
            BITMAP_OUTPUT(0, x, y) = BITMAP_INPUT(0, xx, yy)
            BITMAP_OUTPUT(1, x, y) = BITMAP_INPUT(1, xx, yy)
            BITMAP_OUTPUT(2, x, y) = BITMAP_INPUT(2, xx, yy)
        Next
    Next
    
    '출력 픽처박스 크기를 설정한다.
    Picture1.Move Picture1.Left, Picture1.Top, (w2 + 2) * Screen.TwipsPerPixelX, (h2 + 2) * Screen.TwipsPerPixelY
    
    '비트맵 배열을 이미지로 복사한다.
    Call SetDIBits(Picture1.hdc, Picture1.Image, 0, h2, BITMAP_OUTPUT(0, 0, 0), BITMAP_INFO, DIB_RGB_COLORS)
End Sub

'리셋
Private Sub Command2_Click()
    '이미지 불러오기
    Picture1.Picture = LoadPicture(App.Path + "\image_08.jpg")
End Sub

Private Sub Form_Load()
    '폼을 모니터의 가운데로 이동하기
    Form1.Left = (Screen.Width - Form1.Width) \ 2
    Form1.Top = (Screen.Height - Form1.Height) \ 2

    '이미지 불러오기
    Picture1.Picture = LoadPicture(App.Path + "\image_08.jpg")
End Sub

