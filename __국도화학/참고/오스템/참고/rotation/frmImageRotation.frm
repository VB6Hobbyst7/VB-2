VERSION 5.00
Begin VB.Form frmImageRotation 
   BorderStyle     =   1  '단일 고정
   Caption         =   "그림 회전"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6225
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdRotate 
      Caption         =   "Rotate!"
      Height          =   315
      Left            =   1140
      TabIndex        =   3
      Top             =   4740
      Width           =   915
   End
   Begin VB.TextBox txtDegree 
      Alignment       =   2  '가운데 맞춤
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Text            =   "0"
      Top             =   4740
      Width           =   435
   End
   Begin VB.PictureBox picTarget 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4500
      Left            =   120
      ScaleHeight     =   300
      ScaleMode       =   3  '픽셀
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   120
      Width           =   6000
   End
   Begin VB.PictureBox picImage 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4500
      Left            =   120
      Picture         =   "frmImageRotation.frx":0000
      ScaleHeight     =   300
      ScaleMode       =   3  '픽셀
      ScaleWidth      =   400
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.Label lblUnit 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "deg."
      Height          =   225
      Left            =   720
      TabIndex        =   2
      Top             =   4800
      Width           =   285
   End
End
Attribute VB_Name = "frmImageRotation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32.dll" ( _
    ByVal hDestDC As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal dwRop As Long _
) As Long

Private Declare Function PlgBlt Lib "gdi32.dll" ( _
    ByVal hdcDest As Long, _
    ByRef lpPoint As POINTAPI, _
    ByVal hdcSrc As Long, _
    ByVal nXSrc As Long, _
    ByVal nYSrc As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hbmMask As Long, _
    ByVal xMask As Long, _
    ByVal yMask As Long _
) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const PI As Double = 3.14159265358979

Private Sub cmdRotate_Click()
    
    Dim pts(2) As POINTAPI, CenterX As Double, CenterY As Double, Radian As Double, RotRadian As Double, Distance As Double
    
    ' ### 그림의 중앙 좌표를 구한다.
    CenterX = picTarget.ScaleWidth \ 2
    CenterY = picTarget.ScaleHeight \ 2
    
    RotRadian = CDbl(txtDegree.Text) * PI / 180
    With picTarget
        .Cls ' ### 그림의 버퍼를 깨끗이 비웁니다.
        Distance = Sqr(CenterX ^ 2 + CenterY ^ 2)
        
        Radian = Atn(CenterY / CenterX) + RotRadian
        With pts(0)
            .x = -Math.Cos(Radian) * Distance + CenterX
            .y = -Math.Sin(Radian) * Distance + CenterY
        End With
        
        Radian = Atn(CenterY / (CenterX - picTarget.ScaleWidth)) + RotRadian
        With pts(1)
            .x = Math.Cos(Radian) * Distance + CenterX
            .y = Math.Sin(Radian) * Distance + CenterY
        End With
        
        Radian = Atn((CenterY - picTarget.ScaleHeight) / CenterX) + RotRadian
        With pts(2)
            .x = -Math.Cos(Radian) * Distance + CenterX
            .y = -Math.Sin(Radian) * Distance + CenterY
        End With
        
        PlgBlt picTarget.hDC, pts(0), picImage.hDC, 0&, 0&, picTarget.ScaleWidth, picTarget.ScaleHeight, 0&, 0&, 0&
    End With
    
End Sub

Private Sub Form_Load()
    BitBlt picTarget.hDC, 0&, 0&, picImage.ScaleWidth, picImage.ScaleHeight, picImage.hDC, 0&, 0&, vbSrcCopy
End Sub
