VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBBS822 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "혈액제제마스터"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   Icon            =   "frmBBS822.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   420
      Left            =   6180
      Style           =   1  '그래픽
      TabIndex        =   9
      Top             =   7560
      Width           =   1260
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   420
      Left            =   8820
      Style           =   1  '그래픽
      TabIndex        =   11
      Top             =   7560
      Width           =   1260
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   420
      Left            =   7500
      Style           =   1  '그래픽
      TabIndex        =   10
      Top             =   7560
      Width           =   1260
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      BorderStyle     =   0  '없음
      Height          =   2355
      Left            =   5580
      TabIndex        =   13
      Top             =   4860
      Width           =   4815
      Begin VB.ComboBox cboGroup 
         Height          =   300
         Left            =   810
         Style           =   2  '드롭다운 목록
         TabIndex        =   3
         Top             =   1200
         Width           =   3975
      End
      Begin VB.TextBox txtKeepDay 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   330
         Left            =   810
         ScrollBars      =   2  '수직
         TabIndex        =   4
         Top             =   1560
         Width           =   390
      End
      Begin VB.CheckBox chkPhere 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Pheresis여부"
         Height          =   255
         Left            =   3180
         TabIndex        =   6
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtAbbrNm 
         Height          =   330
         Left            =   810
         ScrollBars      =   2  '수직
         TabIndex        =   2
         Top             =   780
         Width           =   3990
      End
      Begin VB.CheckBox chkExp 
         BackColor       =   &H00DBE6E6&
         Caption         =   "폐기여부"
         Height          =   225
         Left            =   0
         TabIndex        =   7
         Top             =   2040
         Width           =   1110
      End
      Begin VB.TextBox txtCode 
         Height          =   315
         Left            =   810
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "AAAAAAAAAA"
         Top             =   15
         Width           =   1470
      End
      Begin VB.TextBox txtName 
         Height          =   330
         Left            =   810
         ScrollBars      =   2  '수직
         TabIndex        =   1
         Top             =   390
         Width           =   3990
      End
      Begin MedControls1.LisLabel lblExp 
         Height          =   315
         Left            =   2250
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1950
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "폐기일자"
         Appearance      =   0
      End
      Begin MSComCtl2.DTPicker dtpExp 
         Height          =   315
         Left            =   3135
         TabIndex        =   8
         Top             =   1950
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24182787
         CurrentDate     =   36833
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   330
         Left            =   1201
         TabIndex        =   5
         Top             =   1560
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         BuddyControl    =   "txtKeepDay"
         BuddyDispid     =   196614
         OrigLeft        =   1440
         OrigTop         =   1560
         OrigRight       =   1680
         OrigBottom      =   1890
         Max             =   99999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "보관일수"
         Height          =   180
         Left            =   0
         TabIndex        =   23
         Top             =   1650
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "통계그룹"
         Height          =   180
         Left            =   0
         TabIndex        =   22
         Top             =   1230
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "약어"
         Height          =   180
         Left            =   0
         TabIndex        =   21
         Top             =   870
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "명칭"
         Height          =   180
         Left            =   0
         TabIndex        =   20
         Top             =   480
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "코드"
         Height          =   180
         Left            =   0
         TabIndex        =   19
         Top             =   60
         Width           =   360
      End
   End
   Begin MedControls1.LisLabel lblTitle 
      Height          =   375
      Left            =   180
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   180
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   661
      BackColor       =   8421504
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "혈액제제 리스트"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   375
      Left            =   5460
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4320
      Width           =   5070
      _ExtentX        =   8943
      _ExtentY        =   661
      BackColor       =   8421504
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "등록"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   375
      Left            =   180
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4320
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   661
      BackColor       =   8421504
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "검사항목"
      Appearance      =   0
   End
   Begin MSComctlLib.ListView lvwTest 
      Height          =   3375
      Left            =   240
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4740
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14411494
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "검사코드"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "검사항목"
         Object.Width           =   6174
      EndProperty
   End
   Begin FPSpread.vaSpread tblCompo 
      Height          =   3615
      Left            =   240
      TabIndex        =   12
      Tag             =   "10114"
      Top             =   600
      Width           =   10200
      _Version        =   196608
      _ExtentX        =   17992
      _ExtentY        =   6376
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   1
      BorderStyle     =   0
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormulaSync     =   0   'False
      GridShowVert    =   0   'False
      MaxCols         =   10
      MaxRows         =   13
      MoveActiveOnFocus=   0   'False
      OperationMode   =   3
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS822.frx":076A
      StartingColNumber=   2
      VirtualRows     =   24
      VisibleRows     =   13
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   3495
      Index           =   1
      Left            =   195
      Top             =   4680
      Width           =   5175
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   3735
      Index           =   2
      Left            =   195
      Top             =   540
      Width           =   10335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   3495
      Index           =   0
      Left            =   5475
      Top             =   4680
      Width           =   5055
   End
End
Attribute VB_Name = "frmBBS822"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TblColumn
    tcCOMPOCD = 1
    tcCOMPONM
    tcABBRNM
    tcGROUP
    tcKEEPDAY
    tcPHRESIS
    tcEXPDT
    tcGROUPCD
    tcPHRESISCD
End Enum

Private Sub chkExp_Click()
    If chkExp.Value = 1 Then
        lblExp.Visible = True
        dtpExp.Visible = True
        dtpExp = GetSystemDate
    Else
        lblExp.Visible = False
        dtpExp.Visible = False
    End If
End Sub

Private Sub cmdClear_Click()
    Clear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Save = True Then
        Clear
        Query
    End If
End Sub

Private Sub Form_Activate()
'    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    Call ClearAll
    Call QueryGroup
    Call Query
End Sub

Private Sub tblCompo_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim i As Long
    
    If Row = NewRow Then Exit Sub
    
    If NewRow < 0 Then Exit Sub
    
    With tblCompo
        .Row = NewRow
        
        .Col = TblColumn.tcABBRNM:      txtAbbrNm = .Value
        .Col = TblColumn.tcCOMPOCD:     txtCode = .Value
        .Col = TblColumn.tcCOMPONM:     txtName = .Value
        .Col = TblColumn.tcKEEPDAY:     txtKeepDay = .Value
        .Col = TblColumn.tcPHRESISCD:
                                        If .Value = "1" Then
                                            chkPhere.Value = 1
                                        Else
                                            chkPhere.Value = 0
                                        End If
        .Col = TblColumn.tcEXPDT:
                                        If .Value <> "" Then
                                            lblExp.Visible = True
                                            dtpExp.Visible = True
                                            dtpExp = .Value
                                        Else
                                            lblExp.Visible = False
                                            dtpExp.Visible = False
                                        End If
        .Col = TblColumn.tcGROUPCD:
                                        cboGroup.ListIndex = -1
                                        For i = 0 To cboGroup.ListCount - 1
                                            If .Value = medGetP(cboGroup.List(i), 1, " ") Then
                                                cboGroup.ListIndex = i
                                                Exit For
                                            End If
                                        Next i
        Call QueryTest(txtCode)
    End With
End Sub









Private Sub ClearAll()
    medClearTable tblCompo
    tblCompo.MaxRows = 13
    
    Call Clear
End Sub

Private Sub Clear()
    txtCode = ""
    txtName = ""
    txtAbbrNm = ""
    txtKeepDay = ""
    chkPhere.Value = 0
    chkExp.Value = 0
    cboGroup.ListIndex = -1
    lblExp.Visible = False
    dtpExp.Visible = False
    dtpExp = GetSystemDate
End Sub

Private Sub QueryGroup()
    Dim objcom003 As clsCom003
    Dim DrRS As Recordset
    Dim i As Long
    
    Set objcom003 = New clsCom003
    Set DrRS = objcom003.OpenRecordSet(BC2_COMPO_GROUP)
    Set objcom003 = Nothing
    
    cboGroup.Clear
    
    If DrRS Is Nothing Then Exit Sub
    
    With DrRS
        For i = 1 To .RecordCount
            cboGroup.AddItem .Fields("cdval1").Value & "" & " " & .Fields("field1").Value & ""
            .MoveNext
        Next i
    End With
    Set DrRS = Nothing
End Sub

Private Sub QueryTest(ByVal CompoCd As String)
    Dim objCompo As clsComponent
    Dim DrRS As Recordset
    Dim itmX As ListItem
    Dim i As Long
    
    lvwTest.ListItems.Clear
    
    If CompoCd = "" Then Exit Sub
    
    Set objCompo = New clsComponent
    Set DrRS = objCompo.GetTest(CompoCd)
    Set objCompo = Nothing
    
    If DrRS Is Nothing Then Exit Sub
    
    With DrRS
        For i = 1 To .RecordCount
            Set itmX = lvwTest.ListItems.Add()
            itmX.Text = .Fields("testcd").Value & ""
            itmX.SubItems(1) = .Fields("testnm").Value & ""
            .MoveNext
        Next i
    End With
    Set DrRS = Nothing
End Sub

Private Function GetGroupNm(ByVal groupcd As String) As String
    Dim objcom003 As clsCom003
    Dim DrRS As Recordset
    
    Set objcom003 = New clsCom003
    Set DrRS = objcom003.OpenRecordSet(BC2_COMPO_GROUP, groupcd)
    Set objcom003 = Nothing
    
    GetGroupNm = ""
    If DrRS Is Nothing Then Exit Function
    
    With DrRS
        If .RecordCount <= 0 Then
            GetGroupNm = ""
        Else
            GetGroupNm = .Fields("field1").Value & ""
        End If
    End With
    Set DrRS = Nothing
End Function

Private Sub Query()
    Dim objCompo As clsComponent
    Dim DrRS As Recordset
    Dim i As Long

    Set objCompo = New clsComponent
    Set DrRS = objCompo.GetList
    Set objCompo = Nothing

    medClearTable tblCompo
    tblCompo.MaxRows = 27
    
    If DrRS Is Nothing Then Exit Sub

    With tblCompo
        For i = 1 To DrRS.RecordCount
            .Row = i
            If .Row > .MaxRows Then .MaxRows = .MaxRows + 1
            
            .Col = TblColumn.tcABBRNM:    .Value = DrRS.Fields("abbrnm").Value & ""
            .Col = TblColumn.tcCOMPOCD:   .Value = DrRS.Fields("compocd").Value & ""
            .Col = TblColumn.tcCOMPONM:   .Value = DrRS.Fields("componm").Value & ""
            .Col = TblColumn.tcGROUP:     .Value = GetGroupNm(DrRS.Fields("groupcd").Value & "")
            .Col = TblColumn.tcGROUPCD:   .Value = DrRS.Fields("groupcd").Value & ""
            .Col = TblColumn.tcKEEPDAY:   .Value = DrRS.Fields("keepday").Value & ""
            .Col = TblColumn.tcPHRESIS:   .Value = IIf(DrRS.Fields("pherefg").Value & "" = "1", "Y", "")
            .Col = TblColumn.tcPHRESISCD: .Value = DrRS.Fields("pherefg").Value & ""
            .Col = TblColumn.tcEXPDT:
                                          If DrRS.Fields("expdt").Value & "" <> "" Then
                                             .Value = Format(DrRS.Fields("expdt").Value & "", "####-##-##")
                                          Else
                                             .Value = ""
                                          End If
            
            DrRS.MoveNext
        Next i
    End With
    Set DrRS = Nothing
End Sub

Private Function Save() As Boolean
    Dim objCompo As clsComponent

    Set objCompo = New clsComponent
    
    With objCompo
        .abbrnm = txtAbbrNm
        .CompoCd = txtCode
        .componm = txtName
        .keepday = Val(txtKeepDay)
        .pherefg = IIf(chkPhere.Value = 1, "1", "0")
        .groupcd = medGetP(cboGroup.Text, 1, " ")
        If chkExp.Value = 0 Then
            .ExpDt = ""
        Else
            .ExpDt = Format(dtpExp, PRESENTDATE_FORMAT)
        End If
        
        Save = .Save
    End With
    
    Set objCompo = Nothing
End Function





