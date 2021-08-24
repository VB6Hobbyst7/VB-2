VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmBBS821 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "ABO검사항목설정"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   Icon            =   "frmBBS821.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   375
      Left            =   825
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1320
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   661
      BackColor       =   8421504
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Alignment       =   1
      Caption         =   "ABO 검사 항목"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   3075
      Left            =   825
      TabIndex        =   9
      Top             =   1620
      Width           =   8740
      Begin VB.TextBox txtDu 
         Height          =   315
         Left            =   2610
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtABOBack 
         Height          =   315
         Left            =   2610
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtRHSUB 
         Height          =   315
         Left            =   2610
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtABOSUB 
         Height          =   315
         Left            =   2610
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtRh 
         Height          =   315
         Left            =   2610
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtABOFront 
         Height          =   315
         Left            =   2610
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   600
         Width           =   1215
      End
      Begin MedControls1.LisLabel lblABOFront 
         Height          =   315
         Left            =   3900
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   600
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MedControls1.LisLabel lblRH 
         Height          =   315
         Left            =   3900
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1320
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MedControls1.LisLabel lblABOSUB 
         Height          =   315
         Left            =   3900
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1680
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MedControls1.LisLabel lblRHSUB 
         Height          =   315
         Left            =   3900
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2040
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MedControls1.LisLabel lblABOBack 
         Height          =   315
         Left            =   3900
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   960
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MedControls1.LisLabel lblDu 
         Height          =   315
         Left            =   3900
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2400
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label6 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Du 검사항목"
         Height          =   180
         Left            =   1470
         TabIndex        =   28
         Top             =   2460
         Width           =   1005
      End
      Begin VB.Label Label5 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "ABO검사항목(Back Type)"
         Height          =   180
         Left            =   345
         TabIndex        =   20
         Top             =   1020
         Width           =   2160
      End
      Begin VB.Label Label4 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Rh Subgroup 검사항목"
         Height          =   180
         Left            =   630
         TabIndex        =   13
         Top             =   2100
         Width           =   1875
      End
      Begin VB.Label Label3 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "ABO Subtype 검사항목"
         Height          =   180
         Left            =   600
         TabIndex        =   12
         Top             =   1740
         Width           =   1905
      End
      Begin VB.Label Label2 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Rh검사항목"
         Height          =   180
         Left            =   1560
         TabIndex        =   11
         Top             =   1380
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "ABO검사항목(Front Type)"
         Height          =   180
         Left            =   345
         TabIndex        =   10
         Top             =   660
         Width           =   2160
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   480
      Left            =   4725
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   4920
      Width           =   1260
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   480
      Left            =   6045
      Style           =   1  '그래픽
      TabIndex        =   7
      Top             =   4920
      Width           =   1260
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   480
      Left            =   3405
      Style           =   1  '그래픽
      TabIndex        =   8
      Top             =   4920
      Width           =   1260
   End
   Begin VB.Label lblDuO 
      Caption         =   "Label6"
      Height          =   255
      Left            =   7650
      TabIndex        =   29
      Top             =   405
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblRHSUBO 
      Caption         =   "Label6"
      Height          =   255
      Left            =   6390
      TabIndex        =   25
      Top             =   405
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblABOSUBO 
      Caption         =   "Label6"
      Height          =   255
      Left            =   5130
      TabIndex        =   24
      Top             =   405
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblRHO 
      Caption         =   "Label6"
      Height          =   255
      Left            =   3870
      TabIndex        =   23
      Top             =   405
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblABOBackO 
      Caption         =   "Label6"
      Height          =   255
      Left            =   2610
      TabIndex        =   22
      Top             =   405
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblABOFrontO 
      Caption         =   "Label6"
      Height          =   255
      Left            =   1350
      TabIndex        =   21
      Top             =   405
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblAccDt 
      AutoSize        =   -1  'True
      BorderStyle     =   1  '단일 고정
      Caption         =   "####-##-##"
      Height          =   240
      Left            =   6510
      TabIndex        =   18
      Top             =   405
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "frmBBS821"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    Call Query
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Save = True Then
        Call Query
    End If
End Sub

Private Sub Form_Activate()
'    medMain.lblSubMenu.Caption = Me.Caption

End Sub

Private Sub Form_Load()
    Call GetABOList
    Call Clear
    Call Query
End Sub

Private Sub Clear()
    txtABOFront = "": lblABOFront.Caption = "": lblABOFrontO = ""
    txtABOBack = "":  lblABOBack.Caption = "":  lblABOBackO = ""
    txtRh = "":       lblRH.Caption = "":       lblRHO = ""
    txtABOSUB = "":   lblABOSUB.Caption = "":   lblABOSUBO = ""
    txtRHSUB = "":    lblRHSUB.Caption = "":    lblRHSUBO = ""
    txtDu = "":       lblDu.Caption = "":       lblDuO = ""
End Sub

Private Sub Query()
    Dim objcom003 As clsCom003
    Dim DrRS As Recordset
    
    lblAccDt = Format(GetSystemDate, "YYYY-MM-DD")
    
    Set objcom003 = New clsCom003
    Set DrRS = objcom003.OpenRecordSetDay(BC2_ABO_TEST)
    If Not (DrRS Is Nothing) Then
        With DrRS
            If .RecordCount > 0 Then
                lblAccDt = Format(.Fields("cdval1").Value & "", "####-##-##")
                
                txtABOFront = .Fields("field1").Value & "" & "": lblABOFrontO = .Fields("field1").Value & "" & ""
                lblABOFront.Caption = GetTestNm(.Fields("field1").Value & "")
                
                txtABOBack = .Fields("field2").Value & "":  lblABOBackO = .Fields("field2").Value & ""
                lblABOBack.Caption = GetTestNm(.Fields("field2").Value & "")
                
                txtRh = .Fields("field3").Value & "":       lblRHO = .Fields("field3").Value & ""
                lblRH.Caption = GetTestNm(.Fields("field3").Value & "")
                
                txtABOSUB = .Fields("field4").Value & "":   lblABOSUBO = .Fields("field4").Value & ""
                lblABOSUB.Caption = GetTestNm(.Fields("field4").Value & "")
                
                txtRHSUB = .Fields("text1").Value & "":     lblRHSUBO = .Fields("text1").Value & ""
                lblRHSUB.Caption = GetTestNm(.Fields("text1").Value & "")
            
                txtDu = .Fields("text2").Value & "":        lblDuO = .Fields("text2").Value & ""
                lblDu.Caption = GetTestNm(.Fields("text2").Value & "")
            End If
        End With
    End If
    Set DrRS = Nothing
    Set objcom003 = Nothing
End Sub

Private Sub GetABOList()
'    Dim objEditABO As clsBBSMSTStatement
'    Dim DrRS As Recordset
'    Dim i As Long
'    Dim itmX As ListItem
'
'    Set objEditABO = New clsBBSMSTStatement
'    Set DrRS = objEditABO.GetABOList
'    Set objEditABO = Nothing
'
'    lvwABO.ListItems.Clear
'
'    If DrRS Is Nothing Then Exit Sub
'    With DrRS
'        For i = 1 To .RecordCount
'            Set itmX = lvwABO.ListItems.Add()
'            itmX.Text = .Fields("testcd").Value
'            itmX.SubItems(1) = Format(.Fields("applydt").Value, "####-##-##")
'            itmX.SubItems(2) = .Fields("testnm").Value
'            If .Fields("expdt").Value <> "" Then
'                itmX.SubItems(3) = Format(.Fields("expdt").Value, "####-##-##")
'            Else
'                itmX.SubItems(3) = ""
'            End If
'            .MoveNext
'        Next i
'        .RsClose
'    End With
'    Set DrRS = Nothing
End Sub

Private Function Save() As Boolean
    Dim objcom003 As clsCom003
    
    If IsChange = True Then
        Set objcom003 = New clsCom003
        With objcom003
            .CDINDEX = BC2_ABO_TEST
            .cdval1 = Format(GetSystemDate, PRESENTDATE_FORMAT)
            .field1 = txtABOFront
            .field2 = txtABOBack
            .Field3 = txtRh
            .Field4 = txtABOSUB
            .Text1 = txtRHSUB
            .Text2 = txtDu
            
            Save = .Save()
        End With
    Else
        Save = False
    End If
End Function

Private Function IsChange() As Boolean
    If txtABOFront <> lblABOFrontO Then IsChange = True: Exit Function
    If txtABOBack <> lblABOBackO Then IsChange = True:   Exit Function
    If txtRh <> lblRHO Then IsChange = True:             Exit Function
    If txtABOSUB <> lblABOSUBO Then IsChange = True:     Exit Function
    If txtRHSUB <> lblRHSUBO Then IsChange = True:       Exit Function
    If txtDu <> lblDuO Then IsChange = True:       Exit Function
    
    IsChange = False
End Function

Private Function GetTestNm(ByVal TestCd As String) As String
    GetTestNm = GetLabTestNm(TestCd)
End Function

Private Sub txtABOBack_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtABOBack_LostFocus()
    lblABOBack.Caption = GetLabTestNm(txtABOBack, Format(lblAccDt, PRESENTDATE_FORMAT))

    If lblABOBack.Caption = "" Then
        MsgBox "존재하지 않거나, ABO검사항목이 아닙니다.", vbCritical, Me.Caption
        txtABOBack = ""
    End If
End Sub

Private Sub txtABOFront_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtABOFront_LostFocus()

    lblABOFront.Caption = GetLabTestNm(txtABOFront, Format(GetSystemDate, PRESENTDATE_FORMAT))

    If lblABOFront.Caption = "" Then
        MsgBox "존재하지 않거나, ABO검사항목이 아닙니다.", vbCritical, Me.Caption
        txtABOFront = ""
    End If
End Sub

Private Sub txtABOSUB_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtABOSUB_LostFocus()

    lblABOSUB.Caption = GetLabTestNm(txtABOSUB, Format(GetSystemDate, PRESENTDATE_FORMAT))

    If lblABOSUB.Caption = "" Then
        MsgBox "존재하지 않거나, ABO검사항목이 아닙니다.", vbCritical, Me.Caption
        txtABOSUB = ""
    End If
End Sub

Private Sub txtDu_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtDu_LostFocus()

    lblDu.Caption = GetLabTestNm(txtDu, Format(GetSystemDate, PRESENTDATE_FORMAT))

    If lblDu.Caption = "" Then
        MsgBox "존재하지 않거나, ABO검사항목이 아닙니다.", vbCritical, Me.Caption
        txtDu = ""
    End If
End Sub

Private Sub txtRh_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtRh_LostFocus()

    lblRH.Caption = GetLabTestNm(txtRh, Format(GetSystemDate, PRESENTDATE_FORMAT))

    If lblRH.Caption = "" Then
        MsgBox "존재하지 않거나, ABO검사항목이 아닙니다.", vbCritical, Me.Caption
        txtRh = ""
    End If
End Sub

Private Sub txtRHSUB_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtRHSUB_LostFocus()

    lblRHSUB.Caption = GetLabTestNm(txtRHSUB, Format(GetSystemDate, PRESENTDATE_FORMAT))

    If lblRHSUB.Caption = "" Then
        MsgBox "존재하지 않거나, ABO검사항목이 아닙니다.", vbCritical, Me.Caption
        txtRHSUB = ""
    End If
End Sub
