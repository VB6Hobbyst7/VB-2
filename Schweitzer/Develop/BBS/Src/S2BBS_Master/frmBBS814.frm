VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBBS814 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "검체보관일수"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   Icon            =   "frmBBS814.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   420
      Left            =   4080
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   3780
      Width           =   1260
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
      Left            =   5370
      Style           =   1  '그래픽
      TabIndex        =   0
      Tag             =   "128"
      Top             =   3780
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Caption         =   "시간설정"
      Height          =   975
      Left            =   2400
      TabIndex        =   7
      Top             =   2700
      Width           =   5895
      Begin VB.TextBox txtDay 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Left            =   3960
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.CheckBox chkDay 
         BackColor       =   &H00DBE6E6&
         Caption         =   "기타"
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   4
         Top             =   420
         Width           =   675
      End
      Begin VB.CheckBox chkDay 
         BackColor       =   &H00DBE6E6&
         Caption         =   "72Hr"
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   3
         Top             =   420
         Value           =   1  '확인
         Width           =   915
      End
      Begin VB.CheckBox chkDay 
         BackColor       =   &H00DBE6E6&
         Caption         =   "48Hr"
         Height          =   255
         Index           =   1
         Left            =   1260
         TabIndex        =   2
         Top             =   420
         Width           =   975
      End
      Begin VB.CheckBox chkDay 
         BackColor       =   &H00DBE6E6&
         Caption         =   "24Hr"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   420
         Width           =   975
      End
      Begin MSComCtl2.UpDown udDay 
         Height          =   315
         Left            =   4936
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtDay"
         BuddyDispid     =   196613
         OrigLeft        =   2220
         OrigTop         =   1200
         OrigRight       =   2460
         OrigBottom      =   1875
         Max             =   9999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Hr"
         Height          =   180
         Left            =   5220
         TabIndex        =   9
         Top             =   480
         Width           =   180
      End
   End
   Begin VB.Label lblDay 
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblEntDt 
      Height          =   255
      Left            =   7080
      TabIndex        =   10
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmBBS814"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private onPgm As Boolean

Private Sub chkDay_Click(Index As Integer)
    Dim i As Long
    
    If onPgm = True Then Exit Sub
    
    chkDay(Index).Value = 1
    
    onPgm = True
    For i = 0 To 3
        If Index <> i Then
            chkDay(i).Value = 0
        End If
    Next i
    onPgm = False
    
    If Index = 3 Then
        '기타
        txtDay.Enabled = True
        udDay.Enabled = True
        txtDay.BackColor = RGB(255, 255, 255)
    Else
        txtDay = ""
        txtDay.Enabled = False
        udDay.Enabled = False
        txtDay.BackColor = Me.BackColor
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Save = True Then
        Query
    End If
End Sub

Private Sub Form_Activate()
'    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    Call ClearAll
    Call Query
End Sub

Private Sub ClearAll()
    Dim i As Long
    
    onPgm = True
    For i = 0 To 3
        chkDay(i).Value = 0
    Next i
    onPgm = False
    
    txtDay = ""
    txtDay.Enabled = False
    udDay.Enabled = False
    txtDay.BackColor = Me.BackColor
End Sub

Private Sub Query()
    Dim objcom003   As clsCom003
    Dim DrRS        As Recordset
    Dim entdt       As String
    Dim hour        As Long
    Dim i           As Long
    
    Set objcom003 = New clsCom003
    Set DrRS = objcom003.OpenRecordSetDay(BC2_KEEP_HOUR)
    If DrRS Is Nothing Then Exit Sub
    
    entdt = ""
    hour = 0
    With DrRS
        If .RecordCount > 0 Then
            entdt = .Fields("cdval1").Value & ""
            hour = .Fields("field1").Value & ""
        End If
    End With
    
    Set DrRS = Nothing
    Set objcom003 = Nothing
    
    If hour = 24 Then
        chkDay(0).Value = 1
    ElseIf hour = 48 Then
        chkDay(1).Value = 1
    ElseIf hour = 72 Then
        chkDay(2).Value = 1
    Else
        chkDay(3).Value = 1
        txtDay = hour
    End If
End Sub

Private Function Save() As Boolean
    Dim objcom003 As clsCom003
    Dim hour As Long
    
    If chkDay(0).Value = 1 Then
        hour = 24
    ElseIf chkDay(1).Value = 1 Then
        hour = 48
    ElseIf chkDay(2).Value = 1 Then
        hour = 72
    ElseIf chkDay(3).Value = 1 Then
        hour = Val(txtDay)
    Else
        hour = 0
    End If
    
    If hour <= 0 Then
        MsgBox "검체보관시간을 설정하십시요", vbCritical, Me.Caption
        Save = False
        Exit Function
    End If
    
    Set objcom003 = New clsCom003
    objcom003.CDINDEX = BC2_KEEP_HOUR
    objcom003.cdval1 = Format(GetSystemDate, PRESENTDATE_FORMAT)
    objcom003.field1 = hour
    Save = objcom003.Save()
    Set objcom003 = Nothing
End Function





