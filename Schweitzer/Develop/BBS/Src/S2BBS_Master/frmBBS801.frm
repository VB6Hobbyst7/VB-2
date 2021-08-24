VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBBS801 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "임시ID 범위"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   Icon            =   "frmBBS801.frx":0000
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
      Height          =   480
      Left            =   4080
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   4860
      Width           =   1320
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
      Height          =   480
      Left            =   5430
      Style           =   1  '그래픽
      TabIndex        =   5
      Tag             =   "128"
      Top             =   4860
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   2715
      Left            =   2520
      TabIndex        =   0
      Top             =   1860
      Width           =   5775
      Begin VB.TextBox txtLength 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "1"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtPostfix 
         Height          =   315
         Left            =   1980
         TabIndex        =   10
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtPrefix 
         Height          =   315
         Left            =   1980
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtToNo 
         Height          =   315
         Left            =   3540
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtFrNo 
         Height          =   315
         Left            =   1980
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin MedControls1.LisLabel lblFrTo 
         Height          =   315
         Left            =   900
         TabIndex        =   13
         Top             =   2100
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   2955
         TabIndex        =   14
         Top             =   1440
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtLength"
         BuddyDispid     =   196613
         OrigRight       =   240
         OrigBottom      =   735
         Max             =   100
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Length"
         Height          =   180
         Left            =   1290
         TabIndex        =   11
         Top             =   1500
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Postfix"
         Height          =   180
         Left            =   1260
         TabIndex        =   8
         Top             =   780
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Prefix"
         Height          =   180
         Left            =   1380
         TabIndex        =   7
         Top             =   420
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
         Height          =   180
         Left            =   3300
         TabIndex        =   2
         Top             =   1140
         Width           =   135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "임시ID 범위"
         Height          =   180
         Left            =   900
         TabIndex        =   1
         Top             =   1140
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmBBS801"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Trim(txtFrNo.Text) = "" Or Trim(txtToNo.Text) = "" Then
        MsgBox "임시ID의 범위를 입력하여야 합니다", vbCritical, Me.Caption
        Exit Sub
    End If
    
    If Save = True Then
        Query
        txtFrNo.SetFocus
    End If
End Sub

Private Sub Form_Activate()
'    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    ClearAll
    Query
End Sub

Private Sub ClearAll()
    txtFrNo = ""
    txtToNo = ""
End Sub

Private Sub Query()
    Dim objcom003 As clsCom003
    Dim DrRS As Recordset
    
    Set objcom003 = New clsCom003
    Set DrRS = objcom003.OpenRecordSet(BC2_TMP_ID, "0", 1)
    If DrRS Is Nothing Then Exit Sub
    
    With DrRS
        If .RecordCount > 0 Then
            txtFrNo = .Fields("field1").Value & ""
            txtToNo = .Fields("field2").Value & ""
            txtPrefix = .Fields("field3").Value & ""
            txtPostfix = .Fields("field4").Value & ""
            txtLength = .Fields("text1").Value & ""
        End If
    End With
    
    Set DrRS = Nothing
    Set objcom003 = Nothing
End Sub

Private Function Save() As Boolean
    Dim objcom003 As clsCom003
    
    Set objcom003 = New clsCom003
    objcom003.CDINDEX = BC2_TMP_ID
    objcom003.cdval1 = "0"
    objcom003.field1 = Trim(txtFrNo)
    objcom003.field2 = Trim(txtToNo)
    objcom003.Field3 = Trim(txtPrefix)
    objcom003.Field4 = Trim(txtPostfix)
    objcom003.Text1 = Trim(txtLength)
    Save = objcom003.Save()
    Set objcom003 = Nothing
End Function

Private Sub txtFrNo_Change()
    SetFrTo
End Sub

Private Sub txtLength_Change()
    SetFrTo
End Sub

Private Sub txtPostfix_Change()
    SetFrTo
End Sub

Private Sub txtPrefix_Change()
    SetFrTo
End Sub

Private Sub SetFrTo()
    Dim fmt As String
    Dim i As Long
    
    fmt = ""
    For i = 1 To Val(txtLength)
        fmt = fmt & "0"
    Next i
    
    lblFrTo.Caption = txtPrefix & Format(txtFrNo, fmt) & txtPostfix & " ~ " & _
                      txtPrefix & Format(txtToNo, fmt) & txtPostfix
End Sub

Private Sub txtToNo_Change()
    SetFrTo
End Sub




