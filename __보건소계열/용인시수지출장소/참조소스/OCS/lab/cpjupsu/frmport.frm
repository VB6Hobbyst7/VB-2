VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmPort 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   6180
   ClientTop       =   2595
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4485
   Begin VB.OptionButton Option2 
      Caption         =   "Port2"
      Height          =   375
      Left            =   2970
      TabIndex        =   3
      Top             =   810
      Width           =   1230
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Port1"
      Height          =   285
      Left            =   1665
      TabIndex        =   2
      Top             =   855
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   420
      Left            =   1575
      TabIndex        =   1
      Top             =   1485
      Width           =   1860
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   450
      Top             =   810
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Port Test 화면입니다."
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   225
      Width           =   3525
   End
End
Attribute VB_Name = "frmPort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim iPort       As Integer
    Dim iLoop       As Integer
    
    
    iPort = 0
    For iLoop = 1 To 2
        GoSub CHECK_PORT
        If iPort > 0 Then Exit For
    Next
    
    Select Case iPort
        Case 0: MsgBox "Port 를 찾지 못하였습니다"
        Case 1: MsgBox "Com1 "
        Case 2: MsgBox "Com2"
        Case Else: MsgBox " Port 를 찾지 못하였습니다!"
    End Select
    Exit Sub
    
    
    
    
CHECK_PORT:
    Me.MSComm1.CommPort = iLoop
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    
    If MSComm1.PortOpen = True Then
        iPort = iLoop
        Return
    End If
    
    Return
    

End Sub
