VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBBS813 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "혈액 반환 가능 시간"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   Icon            =   "frmBBS813.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   795
      Left            =   2640
      TabIndex        =   2
      Top             =   2640
      Width           =   4935
      Begin VB.TextBox txtHour 
         Height          =   315
         Left            =   2100
         TabIndex        =   4
         Top             =   300
         Width           =   495
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   2596
         TabIndex        =   6
         Top             =   300
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         BuddyControl    =   "txtHour"
         BuddyDispid     =   196611
         OrigRight       =   240
         OrigBottom      =   735
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "분 이내만 가능합니다."
         Height          =   180
         Left            =   2880
         TabIndex        =   7
         Top             =   360
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "혈액 반환은 출고후 "
         Height          =   180
         Left            =   300
         TabIndex        =   5
         Top             =   360
         Width           =   1620
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5130
      Style           =   1  '그래픽
      TabIndex        =   0
      Tag             =   "128"
      Top             =   3720
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   420
      Left            =   3840
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   3720
      Width           =   1260
   End
   Begin VB.Label lblHour 
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   2340
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmBBS813"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private onPgm As Boolean



Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Save = True Then
        Call Query
    End If
End Sub

Private Sub Form_Activate()
'
End Sub

Private Sub Form_Load()
    Call ClearAll
    Call Query
End Sub


Private Sub ClearAll()
    txtHour = ""
    lblHour = ""
End Sub

Private Sub Query()
    Dim objcom003   As clsCom003
    Dim DrRS        As Recordset
    Dim i           As Long
    
    Set objcom003 = New clsCom003
    Set DrRS = objcom003.OpenRecordSetDay(BC2_RETURN_HOUR, "0")
    If DrRS Is Nothing Then Exit Sub
    
    lblHour = -1
    With DrRS
        If .RecordCount > 0 Then
            txtHour = .Fields("field1").Value & ""
            lblHour = .Fields("field1").Value & ""
        End If
    End With
    Set DrRS = Nothing
    Set objcom003 = Nothing
End Sub

Private Function Save() As Boolean
    Dim objcom003 As clsCom003
    
    Set objcom003 = New clsCom003
    objcom003.CDINDEX = BC2_RETURN_HOUR
    objcom003.cdval1 = "0"
    objcom003.field1 = txtHour
    Save = objcom003.Save()
    Set objcom003 = Nothing
End Function





