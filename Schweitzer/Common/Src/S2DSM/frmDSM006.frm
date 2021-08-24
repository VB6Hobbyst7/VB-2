VERSION 5.00
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRctl1.ocx"
Begin VB.Form frmDSM006 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Electronic Signature"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "frmDSM006.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox txtPass 
      DataField       =   "400"
      Height          =   330
      IMEMode         =   3  '사용 못함
      Left            =   3385
      PasswordChar    =   "*"
      TabIndex        =   12
      Top             =   2640
      Width           =   1275
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00DBE6E6&
      Height          =   615
      Left            =   40
      ScaleHeight     =   555
      ScaleWidth      =   4560
      TabIndex        =   6
      Top             =   3000
      Width           =   4620
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00EBF3ED&
         Caption         =   "취소(&C)"
         Height          =   450
         Left            =   2340
         Style           =   1  '그래픽
         TabIndex        =   8
         Top             =   60
         Width           =   1215
      End
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H00EBF3ED&
         Caption         =   "확인(&O)"
         Height          =   450
         Left            =   1080
         Style           =   1  '그래픽
         TabIndex        =   7
         Top             =   60
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1335
      Left            =   40
      TabIndex        =   1
      Top             =   -60
      Width           =   4635
      Begin VB.Label Label5 
         BackStyle       =   0  '투명
         Caption         =   $"frmDSM006.frx":030A
         ForeColor       =   &H004B5BE9&
         Height          =   795
         Left            =   180
         TabIndex        =   10
         Top             =   540
         Width           =   4335
      End
      Begin VB.Label lblDoctNm 
         BackStyle       =   0  '투명
         Caption         =   "테스트"
         Height          =   255
         Left            =   1380
         TabIndex        =   3
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "전자서명의 :"
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00DBE6E6&
      Height          =   1275
      Left            =   40
      ScaleHeight     =   1215
      ScaleWidth      =   4560
      TabIndex        =   0
      Top             =   1320
      Width           =   4620
      Begin DRcontrol1.DrLabel lblNonVerify 
         Height          =   1110
         Left            =   1680
         TabIndex        =   9
         Top             =   60
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   1958
         BackColor       =   -2147483634
         ForeColor       =   4554451
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "궁서체"
            Size            =   26.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "사용불가"
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "이미지확인 :"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "전자 서명 "
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   180
         Width           =   1095
      End
      Begin VB.Image imgSign 
         Appearance      =   0  '평면
         Height          =   1110
         Left            =   1680
         Picture         =   "frmDSM006.frx":03B3
         Stretch         =   -1  'True
         Top             =   60
         Width           =   2805
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   "인증암호 : "
      Height          =   255
      Left            =   2340
      TabIndex        =   11
      Top             =   2700
      Width           =   915
   End
End
Attribute VB_Name = "frmDSM006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

