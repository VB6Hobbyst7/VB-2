VERSION 5.00
Begin VB.Form frmXMRemark 
   BackColor       =   &H00DBE6E6&
   Caption         =   "Patient Remark"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5370
   Icon            =   "frmXMRemark.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3285
   ScaleWidth      =   5370
   StartUpPosition =   1  '소유자 가운데
   Begin VB.TextBox txtRmk 
      Height          =   2355
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   3
      Top             =   120
      Width           =   5115
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   2640
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "15101"
      Top             =   2640
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   510
      Left            =   1320
      Style           =   1  '그래픽
      TabIndex        =   1
      Tag             =   "124"
      Top             =   2640
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   3915
      Style           =   1  '그래픽
      TabIndex        =   0
      Tag             =   "128"
      Top             =   2640
      Width           =   1320
   End
End
Attribute VB_Name = "frmXMRemark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mvarsPtid As String
Private mvarRmk   As String

Public Property Let sPtid(ByVal vData As String)
    mvarsPtid = vData
End Property
Public Property Let rmk(ByVal vData As String)
    mvarRmk = vData
End Property

Private Sub cmdClear_Click()
    txtRmk = ""
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim objSQL As New clsCrossMatching
    
    Call objSQL.SetPtidRmk(mvarsPtid, txtRmk)

    If txtRmk <> "" Then
        frmBBS201.lblrmk.Caption = txtRmk
        frmBBS201.cmdRmk.Caption = "Y"
    Else
        frmBBS201.lblrmk.Caption = ""
        frmBBS201.cmdRmk.Caption = ""
    End If
           
    Set objSQL = Nothing
End Sub

Private Sub Form_Activate()
    txtRmk = mvarRmk
End Sub

