VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frm381EquipMaster 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11100
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   11100
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00F7FFF7&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9345
      MaskColor       =   &H00FFC0C0&
      Style           =   1  '그래픽
      TabIndex        =   32
      Top             =   465
      Width           =   1320
   End
   Begin VB.ListBox lstInstrument 
      BackColor       =   &H00F7FFF7&
      Height          =   6720
      Left            =   210
      TabIndex        =   30
      Top             =   1050
      Width           =   2640
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   6945
      Left            =   2910
      TabIndex        =   16
      Top             =   960
      Width           =   7770
      Begin VB.Frame fraScale 
         BackColor       =   &H00DBE6E6&
         Height          =   1170
         Left            =   3660
         TabIndex        =   39
         Top             =   4770
         Width           =   3885
         Begin VB.OptionButton optTemperature 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Centigrade : C."
            Height          =   240
            Index           =   0
            Left            =   150
            TabIndex        =   11
            Top             =   700
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton optTemperature 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Fahrenheit :  F."
            Height          =   240
            Index           =   1
            Left            =   2040
            TabIndex        =   12
            Top             =   700
            Width           =   1695
         End
         Begin VB.Label Label5 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Temperature Scale"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   210
            TabIndex        =   40
            Top             =   210
            Width           =   2430
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00DBE6E6&
         Height          =   540
         Left            =   1950
         TabIndex        =   38
         Top             =   1260
         Width           =   4530
         Begin VB.OptionButton optEqpDiv 
            BackColor       =   &H00DBE6E6&
            Caption         =   "일반장비 : E."
            Height          =   240
            Index           =   0
            Left            =   165
            TabIndex        =   2
            Top             =   210
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton optEqpDiv 
            BackColor       =   &H00DBE6E6&
            Caption         =   "온도관리장비 :  C."
            Height          =   240
            Index           =   1
            Left            =   2100
            TabIndex        =   3
            Top             =   210
            Width           =   2025
         End
      End
      Begin VB.ComboBox cmbLocation 
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1950
         TabIndex        =   7
         Top             =   3330
         Width           =   2355
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00F4F0F2&
         Caption         =   "저장(&S)"
         Height          =   510
         Left            =   2385
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   6345
         Width           =   1320
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00F4F0F2&
         Caption         =   "화면지움(&C)"
         Height          =   510
         Left            =   5025
         Style           =   1  '그래픽
         TabIndex        =   35
         Top             =   6345
         Width           =   1320
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00F4F0F2&
         Cancel          =   -1  'True
         Caption         =   "종료(&X)"
         Height          =   510
         Left            =   6345
         Style           =   1  '그래픽
         TabIndex        =   34
         Top             =   6330
         Width           =   1320
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00F4F0F2&
         Caption         =   "삭제(&D)"
         Height          =   510
         Left            =   3705
         Style           =   1  '그래픽
         TabIndex        =   33
         Top             =   6345
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker dtpPurDate 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "yyyy""년"" MM""월"" dd""일"""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
         Height          =   375
         Left            =   1950
         TabIndex        =   5
         Top             =   2340
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   64290816
         CurrentDate     =   72937
      End
      Begin VB.ComboBox cmbVendor 
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1950
         TabIndex        =   6
         Top             =   2850
         Width           =   2385
      End
      Begin VB.CheckBox chkDisUse 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Disused Instrument"
         Height          =   330
         Left            =   1950
         TabIndex        =   14
         Top             =   6030
         Width           =   3060
      End
      Begin VB.TextBox txtModelNm 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1950
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1860
         Width           =   2340
      End
      Begin VB.TextBox txtEquipCode 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1950
         MaxLength       =   8
         TabIndex        =   0
         Top             =   420
         Width           =   750
      End
      Begin VB.TextBox txtEquipName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1950
         MaxLength       =   30
         TabIndex        =   1
         Top             =   900
         Width           =   2745
      End
      Begin RichTextLib.RichTextBox rtfRemark 
         Height          =   960
         Left            =   1950
         TabIndex        =   8
         Top             =   3735
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   1693
         _Version        =   393217
         BackColor       =   15857140
         ScrollBars      =   2
         TextRTF         =   $"Lis381.frx":0000
      End
      Begin VB.Frame fraAccept 
         BackColor       =   &H00DBE6E6&
         Height          =   1170
         Left            =   240
         TabIndex        =   20
         Top             =   4770
         Width           =   3345
         Begin VB.TextBox txtLow 
            Alignment       =   1  '오른쪽 맞춤
            BackColor       =   &H00F1F5F4&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2400
            MaxLength       =   5
            TabIndex        =   10
            Top             =   630
            Width           =   735
         End
         Begin VB.TextBox txtHigh 
            Alignment       =   1  '오른쪽 맞춤
            BackColor       =   &H00F1F5F4&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   825
            MaxLength       =   5
            TabIndex        =   9
            Top             =   630
            Width           =   750
         End
         Begin VB.Label Label14 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Low"
            Height          =   330
            Left            =   1815
            TabIndex        =   23
            Top             =   705
            Width           =   540
         End
         Begin VB.Label Label9 
            BackColor       =   &H00DBE6E6&
            Caption         =   "High"
            Height          =   330
            Left            =   195
            TabIndex        =   22
            Top             =   705
            Width           =   540
         End
         Begin VB.Label Label7 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Acceptable Temperature"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   210
            TabIndex        =   21
            Top             =   210
            Width           =   2430
         End
      End
      Begin VB.Label Label2 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Instrument Div :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   37
         Top             =   1365
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Location :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   36
         Top             =   3330
         Width           =   855
      End
      Begin VB.Label lblEquipCd 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Instrument Code :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   29
         Top             =   420
         Width           =   1590
      End
      Begin VB.Label lblEquipNm 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Instrument Name :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   28
         Top             =   900
         Width           =   1695
      End
      Begin VB.Label lblSerialNo 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Model No :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   27
         Top             =   1860
         Width           =   1485
      End
      Begin VB.Label lblVendor 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Vendor :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   26
         Top             =   2895
         Width           =   855
      End
      Begin VB.Label lblPuchaseDate 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Purchased Date :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   25
         Top             =   2385
         Width           =   1590
      End
      Begin VB.Label lblNotes 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Notes :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   24
         Top             =   3735
         Width           =   960
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   855
      Left            =   270
      TabIndex        =   17
      Top             =   7845
      Width           =   10410
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00CDE7FA&
         Caption         =   "Next      >>"
         Height          =   510
         Left            =   5565
         Style           =   1  '그래픽
         TabIndex        =   19
         Top             =   200
         Width           =   1320
      End
      Begin VB.CommandButton cmdPrevious 
         BackColor       =   &H00CDE7FA&
         Caption         =   "<< Previous"
         Height          =   510
         Left            =   4200
         Style           =   1  '그래픽
         TabIndex        =   18
         Top             =   200
         Width           =   1320
      End
   End
   Begin MSComctlLib.TabStrip tabItem 
      Height          =   420
      Left            =   210
      TabIndex        =   31
      Top             =   540
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   " Instrument "
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      Caption         =   "Instrument Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1650
      TabIndex        =   15
      Top             =   600
      Width           =   2625
   End
End
Attribute VB_Name = "frm381EquipMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tInsertData
    sEqpCd As String
    seqpnm As String
    seqpdiv As String
    smodelnm As String
    spurchdt As String
    svandcd As String
    stemplow As String
    stemphigh As String
    sRemark As String
    slocationCd As String
    binusefg As String
    tempscale As String
End Type
Dim cVendor_Cdkey As New Collection
Dim cvendor_Nmkey As New Collection
Dim cLocation_Cdkey As New Collection
Dim cLocation_Nmkey As New Collection
Dim clearFlag  As Boolean

Private Sub cmdAdd_Click()
    Call ClearlstInstrumentSelect
    Call ClearInstrumentInfo
    txtEquipCode.SetFocus
End Sub

Private Sub cmdClear_Click()
    Call ClearlstInstrumentSelect
    Call ClearInstrumentInfo
    optEqpDiv(0).Value = True
    txtEquipCode.SetFocus
End Sub

Private Sub ClearInstrumentInfo()
    txtEquipCode.Text = ""
    txtEquipName.Text = ""
    txtModelNm.Text = ""
    dtpPurDate.Value = Now
    cmbVendor.Text = ""
    cmbLocation.Text = ""
    rtfRemark.Text = ""
    txtHigh.Text = ""
    txtLow.Text = ""
    chkDisUse.Value = 0
    
    optEqpDiv(0).Enabled = True
    optTemperature(0).Enabled = True
    
    fraAccept.Visible = False
    fraScale.Visible = False
End Sub

Private Sub ClearlstInstrumentSelect()
    Dim i%
    clearFlag = True
    With lstInstrument
        For i = 0 To .ListCount - 1
            .Selected(i) = False
        Next i
    End With
    clearFlag = False
End Sub

Private Sub ClearlstInstrumentContent()
    lstInstrument.Clear
End Sub

Private Sub cmdDelete_Click()
    
    Dim sMsg As String
    Dim sRes As Integer, sStyle As Integer
    
    If Trim(txtEquipCode.Text) = "" Then Exit Sub

    sMsg = txtEquipCode.Text & " 에 관한 정보를 모두 삭제합니다" & Chr$(13) & Chr$(10) & _
        "정말 삭제해도 좋습니까?"
    sStyle = vbYesNo + vbCritical + vbDefaultButton2
    
    sRes = MsgBox(sMsg, sStyle, "삭제 확인")
    If sRes = vbYes Then
        Call DeleteEquipInfo
        Call ClearInstrumentInfo
        Call DsplstInstrument
'        medMain.stsBar.Panels(2).Text = "정상적으로 삭제 처리 되었습니다. 다음 작업을 처리하세요"
    Else
        Exit Sub
    End If
    
End Sub
    
Private Sub DeleteEquipInfo()
    
    Dim sSqlDel As String
    Dim objSql As New clsLISSqlStatement
    
On Error GoTo DBExecError
    sSqlDel = objSql.SqlDeleteInstrument(Trim(txtEquipCode.Text))
              
    dbconn.BeginTrans
    dbconn.Execute (sSqlDel)
    dbconn.CommitTrans
    Set objSql = Nothing
'    medMain.stsBar.Panels(2).Text = "정상적으로 삽입 처리 되었습니다. 다음 작업을 처리하세요"
    Exit Sub
DBExecError:
    dbconn.RollbackTrans
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdNext_Click()
    Dim i%
    For i = 0 To lstInstrument.ListCount - 1
        If lstInstrument.Selected(i) = True And (i <> lstInstrument.ListCount - 1) Then
            lstInstrument.Selected(i + 1) = True
            Exit For
        End If
    Next i
    
End Sub

Private Sub cmdPrevious_Click()
    Dim i%
    For i = 0 To lstInstrument.ListCount - 1
        If lstInstrument.Selected(i) = True And i <> 0 Then
            lstInstrument.Selected(i - 1) = True
            Exit For
        End If
    Next i
End Sub

Private Sub cmdSave_Click()
    
    Dim sSqlDel As String
    Dim sSqlInsert As String
    Dim sSqlInsert_New As String
    Dim busefg As Boolean
    Dim spurchdt As String
    Dim vInsertData As tInsertData
    Dim objSql As New clsLISSqlStatement
    
    '-- 온도관리 장비일 경우
    If optEqpDiv(1).Value = True Then
        
        If Trim(txtHigh.Text) = "" Then
            MsgBox "온도를 입력 하세요!", vbCritical, "확인"
            txtHigh.SetFocus
            Exit Sub
        Else
            If IsNumeric(Trim(txtHigh.Text)) = False Then
                MsgBox "숫자형식으로 입력하세요!", vbCritical, "확인"
                txtHigh.SetFocus
                Exit Sub
            End If
        End If
        
        If Trim(txtLow.Text) = "" Then
            MsgBox "온도를 입력 하세요!", vbCritical, "확인"
            txtLow.SetFocus
            Exit Sub
        Else
            If IsNumeric(Trim(txtLow.Text)) = False Then
                MsgBox "숫자형식으로 입력하세요!", vbCritical, "확인"
                txtLow.SetFocus
                Exit Sub
            End If
        End If
        
        If CDbl(txtLow.Text) > CDbl(txtHigh.Text) Then
            MsgBox "온도입력 오류입니다.", vbExclamation
            Exit Sub
        End If
    End If
    
On Error GoTo DBExecError

    sSqlDel = objSql.SqlDeleteInstrument(Trim(txtEquipCode.Text))
    
    dbconn.BeginTrans
    dbconn.Execute (sSqlDel)
    
    vInsertData = MakeInsertData
    
'    If vInsertData.stemphigh <> "Null" Then
'        vInsertData.stemphigh = CInt(vInsertData.stemphigh)
'    End If
'    If vInsertData.stemplow <> "Null" Then
'        vInsertData.stemplow = CInt(vInsertData.stemplow)
'    End If
'    sSqlInsert = objSql.SqlInsertInstrument(vInsertData.seqpcd, vInsertData.seqpnm, vInsertData.seqpdiv, _
'                                            vInsertData.smodelnm, vInsertData.spurchdt, vInsertData.svandcd, _
'                                            vInsertData.stemphigh, vInsertData.stemplow, vInsertData.sRemark, _
'                                            vInsertData.slocationCd, vInsertData.binusefg)
                                            
    sSqlInsert_New = objSql.SqlInsertInstrument_New(vInsertData.sEqpCd, vInsertData.seqpnm, vInsertData.seqpdiv, _
                                                    vInsertData.smodelnm, vInsertData.spurchdt, vInsertData.svandcd, _
                                                    vInsertData.stemphigh, vInsertData.stemplow, vInsertData.sRemark, _
                                                    vInsertData.slocationCd, vInsertData.binusefg, vInsertData.tempscale)
                                            
    dbconn.Execute (sSqlInsert_New)
    dbconn.CommitTrans
'    medMain.stsBar.Panels(2).Text = "정상적으로 삽입 처리 되었습니다. 다음 작업을 처리하세요"

    Call ClearInstrumentInfo
    Call DsplstInstrument
    
    Exit Sub

DBExecError:
    dbconn.RollbackTrans
            
End Sub

Private Function MakeInsertData() As tInsertData
    
    MakeInsertData.sEqpCd = Trim(txtEquipCode.Text)
    
    If Len(Trim(txtEquipName.Text)) <> 0 Then
        MakeInsertData.seqpnm = Trim(txtEquipName.Text)
    Else
        MakeInsertData.seqpnm = "Null"
    End If
    
    If optEqpDiv(1).Value = True Then
        MakeInsertData.seqpdiv = "C"
        
        '-- 추가 Temperature Scale
        If optTemperature(1).Value = True Then
            MakeInsertData.tempscale = "F"
        Else
            MakeInsertData.tempscale = "C"
        End If
    Else
        MakeInsertData.seqpdiv = "E"
        MakeInsertData.tempscale = "Null"
    End If
    
    If Len(Trim(txtModelNm.Text)) <> 0 Then
        MakeInsertData.smodelnm = Trim(txtModelNm.Text)
    Else
        MakeInsertData.smodelnm = "Null"
    End If
    If Len(Format(dtpPurDate.Value, CS_DateDbFormat)) <> 0 Then
        MakeInsertData.spurchdt = Format(dtpPurDate.Value, CS_DateDbFormat)
    Else
        MakeInsertData.spurchdt = "Null"
    End If
    
    If Len(Trim(cmbVendor.Text)) <> 0 Then
        On Error GoTo collectionsearcherrorVendor
        MakeInsertData.svandcd = cvendor_Nmkey.Item(Trim(cmbVendor.Text))
    Else
collectionsearcherrorVendor:
        MakeInsertData.svandcd = Trim(cmbVendor.Text)
    End If
    
    If Len(Trim(cmbLocation.Text)) <> 0 Then
        On Error GoTo collectionsearcherrorLocation
        MakeInsertData.slocationCd = cLocation_Nmkey.Item(Trim(cmbLocation.Text))
    Else
collectionsearcherrorLocation:
        MakeInsertData.slocationCd = Trim(cmbLocation.Text)
    End If
    
    If Len(Trim(txtHigh.Text)) <> 0 Then
        MakeInsertData.stemphigh = Trim(txtHigh.Text)
    Else
        MakeInsertData.stemphigh = "Null"
    End If
    If Len(Trim(txtLow.Text)) <> 0 Then
        MakeInsertData.stemplow = Trim(txtLow.Text)
    Else
        MakeInsertData.stemplow = "Null"
    End If
    If Len(Trim(rtfRemark.Text)) <> 0 Then
        MakeInsertData.sRemark = rtfRemark.Text
    Else
        MakeInsertData.sRemark = "Null"
    End If
    If chkDisUse.Value = 1 Then
        MakeInsertData.binusefg = "0"
    Else
        MakeInsertData.binusefg = "1"
    End If
End Function

Private Sub Form_Load()
   
   Call InitCollection
   Call DsplstInstrument
   Call FillcmbVendor
   If ObjSysInfo.UseBuildingInfo = "1" Then Call FillcmbLocation
   
End Sub

Private Sub FillcmbVendor()
    Dim sSqlGetVendor As String
    Dim rsGetVendor As Recordset
    Dim i%
    Dim objSql As New clsLISSqlStatement
    
    Set rsGetVendor = New Recordset
    rsGetVendor.Open objSql.SqlLAB032CodeList(LC3_Vander, "*"), dbconn
    
    If rsGetVendor.EOF = True Then Exit Sub
    
    For i = 1 To rsGetVendor.RecordCount
        ' 컬렉션에 저장(key = Code)
        cVendor_Cdkey.Add rsGetVendor.Fields("text1").Value, _
                          rsGetVendor.Fields("cdval1").Value
        ' 컬렉션에 저장(key = Name)
        cvendor_Nmkey.Add rsGetVendor.Fields("cdval1").Value, _
                          rsGetVendor.Fields("text1").Value
                    
        ' Vendor Combo에 저장
        cmbVendor.AddItem (rsGetVendor.Fields("text1").Value)
        rsGetVendor.MoveNext
    Next i
    
    Set rsGetVendor = Nothing
    Set objSql = Nothing
    
End Sub

Private Sub InitCollection()
    Dim i%

    For i = 1 To cVendor_Cdkey.Count
        cVendor_Cdkey.Remove (1)
    Next i
    For i = 1 To cvendor_Nmkey.Count
        cvendor_Nmkey.Remove (1)
    Next i
    For i = 1 To cLocation_Cdkey.Count
        cLocation_Cdkey.Remove (1)
    Next i
    For i = 1 To cLocation_Nmkey.Count
        cLocation_Nmkey.Remove (1)
    Next i
    
    optEqpDiv(0).Value = True
    optTemperature(0).Value = True
    
    fraAccept.Visible = False
    fraScale.Visible = False
    
End Sub

Private Sub FillcmbLocation()
    Dim RS As Recordset
    Dim strSQL As String
    
    strSQL = " SELECT cdval1 as buildcd, field1 as buildnm, field2 as buildno " & _
             "   FROM " & T_LAB032 & _
             "  WHERE " & DBW("cdindex", LC3_Buildings, 2)
        
    Set RS = New Recordset
    
    RS.Open strSQL, dbconn
    
    Do Until RS.EOF
        ' 컬렉션에 저장 (key = Code)
        cLocation_Cdkey.Add RS.Fields("buildnm").Value & "", _
                            RS.Fields("buildcd").Value & ""
        ' 컬렉션에 저장 (key = Name)
        cLocation_Nmkey.Add RS.Fields("buildcd").Value & "", _
                            RS.Fields("buildnm").Value & ""
        ' Location Combo에 저장
        cmbLocation.AddItem (RS.Fields("buildnm").Value & "")
        
        RS.MoveNext
    Loop
        
    Set RS = Nothing
    
'    If ObjLISComCode.Building.EOF = True Then Exit Sub
'
'    For i = 1 To ObjLISComCode.Building.RecordCount
'        ' 컬렉션에 저장 (key = Code)
'        cLocation_Cdkey.Add ObjLISComCode.Building.Fields("buildnm"), _
'                            ObjLISComCode.Building.Fields("buildcd")
'        ' 컬렉션에 저장 (key = Name)
'        cLocation_Nmkey.Add ObjLISComCode.Building.Fields("buildcd"), _
'                            ObjLISComCode.Building.Fields("buildnm")
'        ' Location Combo에 저장
'        cmbLocation.AddItem (ObjLISComCode.Building.Fields("buildnm"))
'        ObjLISComCode.Building.MoveNext
'    Next i
        
End Sub

Private Sub DsplstInstrument()
    Dim sSqlGetInstrument As String
    Dim rsGetInstrument As Recordset
    Dim objSql As New clsLISSqlStatement
    Dim i%
    
    Set rsGetInstrument = New Recordset
    rsGetInstrument.Open objSql.SqlInstrument_New, dbconn
    
    If rsGetInstrument.EOF = True Then Exit Sub
    
    lstInstrument.Clear
    For i = 1 To rsGetInstrument.RecordCount
        lstInstrument.AddItem rsGetInstrument.Fields("eqpcd").Value & vbTab & _
                              rsGetInstrument.Fields("eqpnm").Value
                              
        rsGetInstrument.MoveNext
    Next i
    
    Set rsGetInstrument = Nothing
    Set objSql = Nothing
End Sub

Private Sub lstInstrument_Click()
    
    Dim sSqlGetInstrumentInfo As String
    Dim rsGetInstrumentInfo As Recordset
    Dim sEqpCd As String
    Dim objSql As New clsLISSqlStatement
    
    If clearFlag = True Then Exit Sub
    
    sEqpCd = Mid(lstInstrument.Text, 1, _
                 InStr(1, lstInstrument.Text, vbTab, vbTextCompare) - 1)
    
    sSqlGetInstrumentInfo = objSql.SqlInstrument_New(sEqpCd)
    
    Set rsGetInstrumentInfo = New Recordset
    rsGetInstrumentInfo.Open sSqlGetInstrumentInfo, dbconn
    
    txtEquipCode.Text = "" & rsGetInstrumentInfo.Fields("eqpcd").Value
    txtEquipName.Text = "" & rsGetInstrumentInfo.Fields("eqpnm").Value
    
    If rsGetInstrumentInfo.Fields("eqpdiv").Value & "" = "C" Then
        optEqpDiv(1).Value = True
        
        fraAccept.Visible = True
        fraScale.Visible = True
    Else
        optEqpDiv(0).Value = True
        
        fraAccept.Visible = False
        fraScale.Visible = False
    End If
    
    txtModelNm.Text = "" & rsGetInstrumentInfo.Fields("modelnm").Value
    If Trim("" & rsGetInstrumentInfo.Fields("purchdt").Value) <> "" Then
        dtpPurDate.Value = Format("" & rsGetInstrumentInfo.Fields("purchdt").Value, CS_DateMask)
    End If
    Call DspSelVendor("" & rsGetInstrumentInfo.Fields("vandcd").Value)
    Call DspSelLocation("" & rsGetInstrumentInfo.Fields("location").Value)
    rtfRemark.Text = "" & rsGetInstrumentInfo.Fields("remark").Value
    
    txtHigh.Text = Val("" & rsGetInstrumentInfo.Fields("temphigh").Value)
    txtLow.Text = Val("" & rsGetInstrumentInfo.Fields("templow").Value)
    If rsGetInstrumentInfo.Fields("inusefg").Value = "1" Then
        chkDisUse.Value = 0
    Else: chkDisUse.Value = 1
    End If
    
'    If rsGetInstrumentInfo.Fields("tempscale").Value & "" = "F" Then
'        optTemperature(1).Enabled = True
'    Else
'        optTemperature(0).Enabled = True
'    End If
    
    Set rsGetInstrumentInfo = Nothing
    Set objSql = Nothing
End Sub

Private Sub DspSelVendor(vendorCd As String)
    Dim i%
On Error GoTo CollectionSerarchError
    cmbVendor.Text = cVendor_Cdkey.Item(vendorCd)
    Exit Sub
CollectionSerarchError:
    cmbVendor.Text = vendorCd
End Sub

Private Sub DspSelLocation(LocationCd As String)
    Dim i%
On Error GoTo CollectionSerarchError
    cmbLocation.Text = cLocation_Cdkey.Item(LocationCd)
    Exit Sub
CollectionSerarchError:
    cmbLocation.Text = LocationCd
End Sub

Private Sub optEqpDiv_Click(Index As Integer)
    If Index = 1 Then
        fraAccept.Visible = True
        fraScale.Visible = True
    Else
        fraAccept.Visible = False
        fraScale.Visible = False
    End If
End Sub

Private Sub txtHigh_GotFocus()
    With txtHigh
        .SelStart = 0
        .SelLength = Len(txtHigh.Text)
    End With
End Sub

Private Sub txtHigh_KeyPress(KeyAscii As Integer)
'    If ((KeyAscii < vbKey0 Or KeyAscii > vbKey9)) And (vbKeyBack <> KeyAscii) And (KeyAscii <> Asc(".")) Then
'        KeyAscii = 0
'    End If
    
End Sub

Private Sub txtHigh_LostFocus()
    If Trim(txtHigh.Text) <> "" Then
        If IsNumeric(Trim(txtHigh.Text)) = False Then
            MsgBox "수치 입력 오류입니다.", vbExclamation
            txtHigh.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub txtLow_GotFocus()
    With txtLow
        .SelStart = 0
        .SelLength = Len(txtHigh.Text)
    End With
End Sub

Private Sub txtLow_KeyPress(KeyAscii As Integer)
'    If ((KeyAscii < vbKey0 Or KeyAscii > vbKey9)) And (vbKeyBack <> KeyAscii) And (KeyAscii <> Asc(".")) Then
'        KeyAscii = 0
'    End If
End Sub

Private Sub txtLow_LostFocus()
    If Trim(txtLow.Text) <> "" Then
        If IsNumeric(Trim(txtLow.Text)) = False Then
            MsgBox "수치 입력 오류입니다.", vbExclamation
            txtLow.SetFocus
            Exit Sub
        End If
    End If
End Sub
