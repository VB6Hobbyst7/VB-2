VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBBS812 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "XM Step"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   Icon            =   "frmBBS812.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
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
      Left            =   5070
      Style           =   1  '그래픽
      TabIndex        =   0
      Tag             =   "128"
      Top             =   4860
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   420
      Left            =   3780
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   4860
      Width           =   1260
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Caption         =   "교차 시험 단계"
      Height          =   2475
      Left            =   2160
      TabIndex        =   6
      Top             =   1920
      Width           =   5895
      Begin VB.TextBox txtStepNm 
         Height          =   315
         Index           =   3
         Left            =   1680
         TabIndex        =   16
         Top             =   1860
         Width           =   3855
      End
      Begin VB.TextBox txtStepNm 
         Height          =   315
         Index           =   2
         Left            =   1680
         TabIndex        =   14
         Top             =   1500
         Width           =   3855
      End
      Begin VB.TextBox txtStepNm 
         Height          =   315
         Index           =   1
         Left            =   1680
         TabIndex        =   12
         Top             =   1140
         Width           =   3855
      End
      Begin VB.TextBox txtStepNm 
         Height          =   315
         Index           =   0
         Left            =   1680
         TabIndex        =   10
         Top             =   780
         Width           =   3855
      End
      Begin VB.CheckBox chkStep 
         BackColor       =   &H00DBE6E6&
         Caption         =   "3 Step"
         Height          =   255
         Index           =   0
         Left            =   540
         TabIndex        =   1
         Top             =   420
         Width           =   825
      End
      Begin VB.CheckBox chkStep 
         BackColor       =   &H00DBE6E6&
         Caption         =   "4 Step"
         Height          =   255
         Index           =   1
         Left            =   1500
         TabIndex        =   2
         Top             =   420
         Value           =   1  '확인
         Width           =   825
      End
      Begin VB.CheckBox chkStep 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Etc"
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   3
         Top             =   435
         Width           =   675
      End
      Begin VB.TextBox txtStep 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   315
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   315
      End
      Begin MSComCtl2.UpDown udStep 
         Height          =   315
         Left            =   3435
         TabIndex        =   8
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtStep"
         BuddyDispid     =   196615
         OrigRight       =   240
         OrigBottom      =   735
         Max             =   4
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   315
         Left            =   540
         TabIndex        =   9
         Top             =   780
         Width           =   1155
         _ExtentX        =   2037
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
         Caption         =   "1 Step"
      End
      Begin MedControls1.LisLabel LisLabel2 
         Height          =   315
         Left            =   540
         TabIndex        =   11
         Top             =   1140
         Width           =   1155
         _ExtentX        =   2037
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
         Caption         =   "2 Step"
      End
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   315
         Left            =   540
         TabIndex        =   13
         Top             =   1500
         Width           =   1155
         _ExtentX        =   2037
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
         Caption         =   "3 Step"
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Left            =   540
         TabIndex        =   15
         Top             =   1860
         Width           =   1155
         _ExtentX        =   2037
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
         Caption         =   "4 Step"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Step"
         Height          =   180
         Left            =   3720
         TabIndex        =   7
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.Label lblStepCnt 
      Height          =   255
      Left            =   2160
      TabIndex        =   18
      Top             =   1620
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblEntDt 
      Height          =   255
      Left            =   6780
      TabIndex        =   17
      Top             =   1620
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmBBS812"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private onPgm As Boolean


Private Sub chkStep_Click(Index As Integer)
    Dim i As Long
    Dim stepcnt As Long
    
    If onPgm = True Then Exit Sub
    
    chkStep(Index).Value = 1
    
    onPgm = True
    For i = 0 To 2
        If Index <> i Then
            chkStep(i).Value = 0
        End If
    Next i
    onPgm = False
    
    Select Case Index
        Case 0:
            txtStep = ""
            txtStep.BackColor = Me.BackColor
            udStep.Enabled = False
        Case 1:
            txtStep = ""
            txtStep.BackColor = Me.BackColor
            udStep.Enabled = False
        Case 2:
            txtStep.BackColor = RGB(255, 255, 255)
            udStep.Enabled = True
            txtStep.Text = "2"
    End Select
    
    MakeTblStep
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
    ClearAll
    Query
End Sub

Private Sub ClearAll()
    Dim i As Long
    
    onPgm = True
    For i = 0 To 2
        chkStep(i).Value = 0
    Next i
    onPgm = False
    
    txtStep = ""
    txtStep.BackColor = Me.BackColor
    udStep.Enabled = False

End Sub

Private Sub Query()
    Dim objcom003 As clsCom003
    Dim DrRS As Recordset
    Dim entdt As String
    Dim stepcnt As String
    Dim stepname As String
    Dim i As Long
    
    lblEntDt = Format(GetSystemDate, "YYYY-MM-DD")
    
    Set objcom003 = New clsCom003
    Set DrRS = objcom003.OpenRecordSetDay(BC2_XM_STEP)
    If DrRS Is Nothing Then Exit Sub
    
    With DrRS
        If .RecordCount > 0 Then
            entdt = .Fields("cdval1").Value & ""
            stepcnt = .Fields("field1").Value & ""
            stepname = .Fields("text1").Value & ""
        End If
    End With
    
    Set DrRS = Nothing
    Set objcom003 = Nothing
    
    If entdt = "" Then Exit Sub
    
    lblEntDt = Format(entdt, "####-##-##")
    
    If stepcnt = 3 Then
        chkStep(0).Value = 1
    ElseIf stepcnt = 4 Then
        chkStep(1).Value = 1
    Else
        chkStep(2).Value = 1
        txtStep = stepcnt
    End If
    
    lblStepCnt = stepcnt
    For i = 1 To stepcnt
        txtStepNm(i - 1) = medGetP(stepname, i, ";")
    Next i
    
    Call MakeTblStep
End Sub

Private Function Save() As Boolean
    Dim objcom003 As clsCom003
    Dim stepcnt As Long
    Dim stepname As String
    Dim i As Long
    
    If chkStep(0).Value = 1 Then
        stepcnt = 3
    ElseIf chkStep(1).Value = 1 Then
        stepcnt = 4
    ElseIf chkStep(2).Value = 1 Then
        stepcnt = Val(txtStep)
    Else
        stepcnt = -1
    End If
    
    
    stepname = ""
    For i = 1 To stepcnt
        If Trim(txtStepNm(i - 1)) = "" Then
            MsgBox "STEP의 명칭을 모두 입력하여야합니다", vbCritical, Me.Caption
            Save = False
            Exit Function
        End If
        If stepname = "" Then
            stepname = txtStepNm(i - 1)
        Else
            stepname = stepname & ";" & txtStepNm(i - 1)
        End If
    Next i
    
    
    Set objcom003 = New clsCom003
    objcom003.CDINDEX = BC2_XM_STEP
    If stepcnt <> Val(lblStepCnt) Then
        objcom003.cdval1 = Format(GetSystemDate, PRESENTDATE_FORMAT)
    Else
        objcom003.cdval1 = Format(GetSystemDate, PRESENTDATE_FORMAT)
    End If
    objcom003.field1 = stepcnt
    objcom003.Text1 = stepname
    Save = objcom003.Save()
    Set objcom003 = Nothing
End Function

Private Sub MakeTblStep()
    Dim step As Long
    
    If chkStep(0).Value = 1 Then
        step = 3
    ElseIf chkStep(1).Value = 1 Then
        step = 4
    Else
        step = txtStep
    End If
    
    Select Case step
        Case 1:
            txtStepNm(0).Enabled = True
            txtStepNm(1).Enabled = False
            txtStepNm(2).Enabled = False
            txtStepNm(3).Enabled = False
        Case 2:
            txtStepNm(0).Enabled = True
            txtStepNm(1).Enabled = True
            txtStepNm(2).Enabled = False
            txtStepNm(3).Enabled = False
        Case 3:
            txtStepNm(0).Enabled = True
            txtStepNm(1).Enabled = True
            txtStepNm(2).Enabled = True
            txtStepNm(3).Enabled = False
        Case 4:
            txtStepNm(0).Enabled = True
            txtStepNm(1).Enabled = True
            txtStepNm(2).Enabled = True
            txtStepNm(3).Enabled = True
    End Select
End Sub

Private Sub txtStep_Change()
    If Trim(txtStep) = "" Then Exit Sub
    MakeTblStep
End Sub




