VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS309 
   BackColor       =   &H00DBE6E6&
   Caption         =   "혈액 Transfer"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14535
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   14535
   WindowState     =   2  '최대화
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   16
      Tag             =   "128"
      Top             =   8430
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel11 
      Height          =   315
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   556
      BackColor       =   8388608
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
      Caption         =   "혈액 Transfer"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   8040
      Left            =   75
      TabIndex        =   1
      Top             =   300
      Width           =   7950
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   3
         Left            =   75
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   255
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Center"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   4
         Left            =   75
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   630
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "혈액 번호"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   6
         Left            =   75
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1005
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Component"
         Appearance      =   0
      End
      Begin VB.ListBox lstCompo 
         Appearance      =   0  '평면
         Height          =   1830
         Left            =   1290
         TabIndex        =   7
         Top             =   1335
         Visible         =   0   'False
         Width           =   4170
      End
      Begin VB.TextBox txtCenterCd 
         Appearance      =   0  '평면
         Height          =   315
         Left            =   1290
         TabIndex        =   9
         Top             =   255
         Width           =   750
      End
      Begin VB.CommandButton cmdCenterList 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   8
         Top             =   255
         Width           =   330
      End
      Begin VB.TextBox txtBldNo 
         Appearance      =   0  '평면
         Height          =   315
         Left            =   1290
         TabIndex        =   6
         Top             =   630
         Width           =   2070
      End
      Begin VB.CheckBox chkBar 
         BackColor       =   &H00DBE6E6&
         Caption         =   "바코드입력"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3990
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   735
         Width           =   1575
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00F4F0F2&
         Caption         =   "화면지움(&C)"
         Height          =   510
         Left            =   6525
         Style           =   1  '그래픽
         TabIndex        =   4
         Tag             =   "124"
         Top             =   7230
         Width           =   1320
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00F4F0F2&
         Caption         =   "실행(&S)"
         Height          =   510
         Left            =   5205
         Style           =   1  '그래픽
         TabIndex        =   3
         Tag             =   "15101"
         Top             =   7230
         Width           =   1320
      End
      Begin MedControls1.LisLabel lblCenterNm 
         Height          =   315
         Left            =   2385
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   255
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   556
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblCompoCd 
         Height          =   315
         Left            =   1290
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1005
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblCompoNm 
         Height          =   315
         Left            =   2385
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1005
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   556
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblTotCnt 
         Height          =   315
         Left            =   6750
         TabIndex        =   13
         Top             =   1005
         Width           =   1095
         _ExtentX        =   1931
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
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin FPSpread.vaSpread tblResult 
         Height          =   5640
         Left            =   90
         TabIndex        =   2
         Top             =   1470
         Width           =   7770
         _Version        =   196608
         _ExtentX        =   13705
         _ExtentY        =   9948
         _StockProps     =   64
         BackColorStyle  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14411494
         GridShowVert    =   0   'False
         MaxCols         =   9
         MaxRows         =   50
         ScrollBars      =   2
         SelectBlockOptions=   2
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS309.frx":0000
         TextTip         =   4
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   1
         Left            =   5535
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1005
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Total Unit"
         Appearance      =   0
      End
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   8040
      TabIndex        =   14
      Top             =   45
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   556
      BackColor       =   8388608
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
      Caption         =   "이동혈액조회"
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   8040
      Left            =   8040
      TabIndex        =   15
      Top             =   300
      Width           =   6420
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00F4F0F2&
         Caption         =   "조회(&Q)"
         Height          =   510
         Left            =   4875
         Style           =   1  '그래픽
         TabIndex        =   21
         Tag             =   "124"
         Top             =   570
         Width           =   1320
      End
      Begin VB.ComboBox cboCenter 
         Height          =   300
         ItemData        =   "frmBBS309.frx":077F
         Left            =   1290
         List            =   "frmBBS309.frx":0781
         Style           =   2  '드롭다운 목록
         TabIndex        =   17
         Top             =   375
         Width           =   1260
      End
      Begin MSComCtl2.DTPicker dtpFMonth 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "gg yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   1290
         TabIndex        =   18
         Top             =   750
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   59899907
         CurrentDate     =   36799
      End
      Begin MSComCtl2.DTPicker dtpTMonth 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "gg yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   2745
         TabIndex        =   19
         Top             =   750
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   59899907
         CurrentDate     =   36799
      End
      Begin FPSpread.vaSpread tbldata 
         Height          =   5625
         Left            =   45
         TabIndex        =   22
         Top             =   1485
         Width           =   6315
         _Version        =   196608
         _ExtentX        =   11139
         _ExtentY        =   9922
         _StockProps     =   64
         BackColorStyle  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14411494
         GridShowVert    =   0   'False
         MaxCols         =   7
         MaxRows         =   50
         ScrollBars      =   2
         SelectBlockOptions=   2
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS309.frx":0783
         TextTip         =   4
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   375
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Center"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   2
         Left            =   60
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   750
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "조회기간 "
         Appearance      =   0
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
         Height          =   180
         Left            =   2580
         TabIndex        =   20
         Top             =   825
         Width           =   135
      End
   End
End
Attribute VB_Name = "frmBBS309"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private objSql As clsBBSSQLStatement
'Private WithEvents mnuPopup As Menu
'Private WithEvents mnuDelete As Menu
Private WithEvents objPop As clsPopupMenu
Attribute objPop.VB_VarHelpID = -1
Private Const MENU_DEL& = 1
Private lngRow As Long
Private WithEvents objMyList As clsPopUpList
Attribute objMyList.VB_VarHelpID = -1

Private Sub cboCenter_Click()
    tbldata.MaxRows = 0
End Sub

Private Sub cmdCenterList_Click()

    Set objSql = New clsBBSSQLStatement
    Set objMyList = New clsPopUpList
    
    txtCenterCd.Text = "": lblCenterNm.Caption = ""
    
    With objMyList
'        .BackColor = Me.BackColor
        .Connection = DBConn
        .FormCaption = "센터조회": .ColumnHeaderText = "코드;코드명"
'        .Width = .Width + 300: .ColSize(0) = 1000
        Call .LoadPopUp(objSql.GetCenterNm) ', 2350, 7650)
        If .SelectedString <> "" Then
            txtCenterCd.Text = medGetP(.SelectedString, 1, ";")
            lblCenterNm.Caption = medGetP(.SelectedString, 2, ";")
        End If
    
    End With
    
    If txtCenterCd.Text = ObjSysInfo.BuildingCd Then
        MsgBox "동일한 혈액센타를 선택 할 수 없습니다.", vbInformation, Me.Caption
        txtCenterCd.Text = "": lblCenterNm.Caption = ""
        txtCenterCd.SetFocus
    End If

    Set objSql = Nothing
    Set objMyList = Nothing
    

End Sub

Private Sub cmdClear_Click()
    Clear
    
End Sub

Private Sub cmdExit_Click()
    If Not objSql Is Nothing Then
        Set objSql = Nothing
    End If
    Unload Me
End Sub

Private Sub cmdQuery_Click()
    Dim objBg    As New clsBeginTrans
    Dim RS       As Recordset
    Dim SSQL     As String
    Dim CenterCd As String
    Dim fDate    As String
    Dim tDate    As String
    
    Dim ii      As Integer
    
    CenterCd = medGetP(cboCenter.Text, 1, " ")
    
    fDate = Format(dtpFMonth.value, "YYYYMMDD")
    tDate = Format(dtpTMonth.value, "YYYYMMDD")
    
    Me.MousePointer = 11
    tbldata.MaxRows = 0
    SSQL = objBg.TransBloodSQL(CenterCd, fDate, tDate)
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        With tbldata
            Do Until RS.EOF
                If .DataRowCnt + 1 > .MaxRows Then
                    .MaxRows = .MaxRows + 1
                End If
                .Row = .DataRowCnt + 1
                .Col = 1: .value = Format(RS.Fields("transdt").value & "", "####-##-##")
                .Col = 2: .value = RS.Fields("bldsrc").value & "" & "-" & RS.Fields("bldyy").value & "" & "-" & _
                                   Format(RS.Fields("bldno").value & "", "000000")
                .Col = 3: .value = RS.Fields("abo").value & "" & RS.Fields("rh").value & ""
                .Col = 4: .value = RS.Fields("abbrnm").value & ""
                .Col = 5: .value = RS.Fields("volume").value & ""
                
'                ObjBBSComCode.Building.KeyChange (Trim("" & Rs.Fields("centercd").value & ""))
                .Col = 6: .value = GetBuildNm(Trim("" & RS.Fields("centercd").value & "")) 'ObjBBSComCode.Building.Fields("buildnm")
'                ObjBBSComCode.Building.KeyChange (Trim("" & Rs.Fields("buildcd").value & ""))
                .Col = 7: .value = GetBuildNm(Trim("" & RS.Fields("buildcd").value & "")) 'ObjBBSComCode.Building.Fields("buildnm")
                RS.MoveNext
            Loop
        End With
    Else
        MsgBox "혈액이동정보가 없습니다.", vbInformation + vbOKOnly, "Info"
    End If
    Set RS = Nothing
    Set objBg = Nothing
    Me.MousePointer = 0
    
End Sub

Private Sub cmdSave_Click()
    Dim objBg      As clsBeginTrans
    Dim strBldno   As String
    Dim strBldSrc  As String
    Dim strBldYY   As String
    Dim lngBldNo   As Long
    Dim strCompocd As String
    
    Dim sVOL       As String
    Dim sABO       As String
    Dim sRH        As String
    
    Dim SSQL       As String
    Dim i          As Long
    
    If txtCenterCd.Text = "" Then
        MsgBox "혈액센타코드를 넣어 주세요.", vbInformation, Me.Caption
        txtCenterCd.SetFocus
        Exit Sub
    End If
On Error GoTo SAVE_ERROR
    
    DBConn.BeginTrans
    Set objBg = New clsBeginTrans
    
    With tblResult
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = 1: strBldno = .value
            .Col = 9: strCompocd = .value
            .Col = 2: sABO = Mid(.value, 1, Len(.value) - 1)
                      sRH = Mid(.value, Len(.value))
            .Col = 4: sVOL = .value
            
            strBldSrc = Mid(strBldno, 1, 2)
            strBldYY = Mid(strBldno, 4, 2)
            lngBldNo = Val(Mid(strBldno, 7))
            SSQL = objBg.SetBldStorageUpdateByCenterCd(strBldSrc, strBldYY, lngBldNo, _
                                                        strCompocd, Trim(txtCenterCd.Text))
            DBConn.Execute SSQL
            
            SSQL = objBg.SetTransDelteSSQL(strBldSrc, strBldYY, CStr(lngBldNo), strCompocd, ObjSysInfo.BuildingCd)
            DBConn.Execute SSQL
            SSQL = objBg.SetTransInfoSSQL(strBldSrc, strBldYY, CStr(lngBldNo), strCompocd, Trim(txtCenterCd.Text), _
                                        sABO, sRH, sVOL)
            DBConn.Execute SSQL
            
        Next
    End With
    
    DBConn.CommitTrans
    Call Clear
    MsgBox "변경하였습니다.", vbInformation + vbOKOnly, "혈액 Transfer"
    Set objBg = Nothing
    Exit Sub
    
SAVE_ERROR:
    DBConn.RollbackTrans
    Set objBg = Nothing
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub dtpFMonth_Change()
    tbldata.MaxRows = 0
End Sub

Private Sub dtpTMonth_Change()
    tbldata.MaxRows = 0
End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    Call SetCenterCombo
    dtpFMonth.value = GetSystemDate
    dtpTMonth.value = GetSystemDate
End Sub
Private Sub SetCenterCombo()
    Dim objcom003 As clsCom003
    Dim i As Long
    
    Set objcom003 = New clsCom003
    Call objcom003.AddComboBox(BC2_CENTER, cboCenter, True)
    Set objcom003 = Nothing
    
    cboCenter.ListIndex = -1
    
    For i = 0 To cboCenter.ListCount - 1
        If ObjSysInfo.BuildingCd = medGetP(cboCenter.List(i), 1, " ") Then
            cboCenter.ListIndex = i
            Exit For
        End If
    Next i
End Sub
Private Sub lstCompo_Click()
    With lstCompo
        lblCompoCd.Caption = medGetP(.List(.ListIndex), 1, vbTab)
        lblCompoNm.Caption = medGetP(.List(.ListIndex), 2, vbTab)
    End With
    lstCompo.Visible = False
    Search
End Sub

Private Sub objPop_Click(ByVal vMenuID As Long)
    Select Case vMenuID
        Case MENU_DEL
            With tblResult
                .Row = lngRow
                .Col = -1
                .Action = 5
                .MaxRows = .MaxRows - 1
                lblTotCnt.Caption = .MaxRows
            End With
    End Select
End Sub

Private Sub txtBldNo_Change()
    If chkBar.value = 1 Then Exit Sub
    Dim lngLen As Long
    
    With txtBldNo
        lngLen = Len(Trim(.Text))
        If lngLen = 2 Then
                .Text = .Text & "-"
                .SelStart = Len(.Text)
        End If
        If lngLen > 2 And lngLen = 5 Then
            .Text = .Text & "-"
            .SelStart = Len(.Text)
        End If
    End With
End Sub

Private Sub txtBldNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txtBldNo_KeyPress(KeyAscii As Integer)
    If chkBar.value = 1 Then Exit Sub
    
    If Len(txtBldNo) <> 3 Or Len(txtBldNo) <> 6 Then
        If KeyAscii = vbKeyInsert Then KeyAscii = 0
    End If
    
    If KeyAscii = vbKeyBack Then
        With txtBldNo
            If .Text = "" Then Exit Sub
            If Mid(.Text, Len(.Text)) = "-" Then
                .Text = Mid(.Text, 1, Len(.Text) - 2)
                .SelStart = Len(.Text)
                KeyAscii = 0
            End If
        End With
    End If
End Sub


Private Sub txtBldNo_LostFocus()
    Dim RS          As Recordset
    Dim rs1         As Recordset
    Dim strSQL      As String
    Dim strSql1     As String
    Dim strChkBldNo As String
    Dim strBldno    As String
    Dim strBldSrc   As String
    Dim strBldYY    As String
    Dim lngBldNo    As Long
    Dim i As Long
    
    If Trim(txtBldNo) = "" Then Exit Sub
    

    If chkBar.value = 1 Then
        If Len(txtBldNo.Text) < 5 Then Exit Sub
        strBldno = Mid(txtBldNo, 1, 2) & "-" & _
                   Mid(txtBldNo, 3, 2) & "-" & _
                   Mid(txtBldNo, 5, 6)
    Else
        strBldno = Mid(txtBldNo, 1, 6) & Format(Mid(txtBldNo, 7), "00000#")
        
    End If
    
    strBldSrc = Mid(strBldno, 1, 2)
    strBldYY = Mid(strBldno, 4, 2)
    lngBldNo = Val(Mid(strBldno, 7))
    
    Set objSql = New clsBBSSQLStatement
    
    strSQL = objSql.GetCompoCd(strBldSrc, strBldYY, lngBldNo)
'    objSql.setDbConn DBConn
    
    Set RS = New Recordset
    RS.Open strSQL, DBConn
    If RS.EOF = False Then
        If RS.RecordCount = 1 Then
            strSql1 = objSql.GetStorageHistory(strBldSrc, strBldYY, lngBldNo, RS.Fields("compocd").value & "", ObjSysInfo.BuildingCd)
            Set rs1 = Nothing
            Set rs1 = New Recordset
            rs1.Open strSql1, DBConn
            
            If rs1.EOF = False Then
                lblCompoCd.Caption = RS.Fields("compocd").value & ""
                lblCompoNm.Caption = RS.Fields("field1").value & ""
                Search
            Else
                MsgBox "혈액센터 안에 존재하지 않은 혈액입니다.", vbInformation, Me.Caption
                txtBldNo.Text = ""
                Set rs1 = Nothing
                Exit Sub
            End If
        Else
            With lstCompo
                .Clear
                Do Until RS.EOF
                    strSql1 = objSql.GetStorageHistory(strBldSrc, strBldYY, lngBldNo, RS.Fields("compocd").value & "", ObjSysInfo.BuildingCd)
                    Set rs1 = Nothing
                    Set rs1 = New Recordset
                    rs1.Open strSql1, DBConn
                    If rs1.EOF = False Then
                        If rs1.Fields("reserved").value & "" <> 1 And rs1.Fields("autofg").value & "" <> 1 Then
                            .AddItem RS.Fields("compocd").value & "" & vbTab & RS.Fields("field1").value & ""
                        End If
                    End If
                    Set rs1 = Nothing
                    RS.MoveNext
                Loop
                If .ListCount = 1 Then
                    lblCompoCd.Caption = medGetP(.List(0), 1, vbTab)
                    lblCompoNm.Caption = medGetP(.List(0), 2, vbTab)
                    Search
                ElseIf .ListCount > 1 Then
                    .Visible = True
                Else
                    MsgBox "혈액센터 안에 존재하지 않은 혈액입니다.", vbInformation, Me.Caption
                    txtBldNo.Text = ""
                    lblCompoCd.Caption = ""
                    lblCompoNm.Caption = ""
                End If
            End With
        End If
    Else
        MsgBox "혈액센터 안에 존재하지 않은 혈액입니다.", vbInformation, Me.Caption
        txtBldNo.Text = ""
        lblCompoCd.Caption = ""
        lblCompoNm.Caption = ""
    End If
    txtBldNo.Text = "": lblCompoCd.Caption = "": lblCompoNm.Caption = ""
    txtBldNo.SetFocus
    Set RS = Nothing
    Set objSql = Nothing
End Sub

Private Sub Search()
    Dim RS As New Recordset
    Dim rs1 As New Recordset
    Dim totslot As Long
    Dim strSQL As String
    Dim strSql1 As String
    Dim strChkBldNo As String
    Dim strBldno As String
    Dim strBldSrc As String
    Dim strBldYY As String
    Dim lngBldNo As Long
    Dim i As Long
    
    If Trim(txtBldNo) = "" Then Exit Sub
'
'    For i = 1 To Len(txtBldNo)
'        If IsNumeric(Mid(txtBldNo, i, 1)) Then
'            strChkBldNo = strChkBldNo & Mid(txtBldNo, i, 1)
'        End If
'    Next i
'
'    If strChkBldNo = "" Then Exit Sub
    
    If chkBar.value = 1 Then
        strBldno = Mid(txtBldNo, 1, 2) & "-" & _
                   Mid(txtBldNo, 3, 2) & "-" & _
                   Format(Mid(Mid(txtBldNo, 5), 1, Len(Mid(txtBldNo, 5)) - 2), "00000#")
    Else
        strBldno = Mid(txtBldNo, 1, 6) & Format(Mid(txtBldNo, 7), "00000#")
        
    End If
    
    
'    strBldNo = Mid(strChkBldNo, 1, 2) & "-" & Mid(strChkBldNo, 3, 2) & "-" & Format(Mid(strChkBldNo, 5), "00000#")
    
    strBldSrc = Mid(strBldno, 1, 2)
    strBldYY = Mid(strBldno, 4, 2)
    lngBldNo = Val(Mid(strBldno, 7))
    
    Set objSql = New clsBBSSQLStatement
    strSQL = objSql.GetStorageHistory(strBldSrc, strBldYY, lngBldNo, Trim(lblCompoCd.Caption))
'    objSql.setDbConn DBConn
    Set RS = New Recordset
    RS.Open strSQL, DBConn
    If RS.EOF = True And RS.BOF = True Then
        lblCompoCd.Caption = ""
        lblCompoNm.Caption = ""
        txtBldNo.SetFocus
        Set RS = Nothing
        Set objSql = Nothing
        Exit Sub
    Else
        '스프레드에 내용을 뿌리자...
        With tblResult
            .MaxRows = .DataRowCnt
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 1: .value = strBldno
            .Col = 2: .value = Trim(RS.Fields("abo").value & "") & Trim(RS.Fields("rh").value & "")
            
            strSql1 = objSql.GetCompCdForCboBox(Trim(RS.Fields("compocd").value & ""))
            
            Set rs1 = New Recordset
            rs1.Open strSql1, DBConn
            If rs1.EOF = False Then
                .Col = 3: .value = rs1.Fields("field1").value & ""
            End If
            Set rs1 = Nothing
            .Col = 4: .value = Trim(RS.Fields("volumn").value & "")
            .Col = 5: .value = Format(Trim(RS.Fields("coldt").value & ""), "####-##-##")
            .Col = 6: .value = Trim(RS.Fields("available").value & "")
            .Col = 7: .value = Format(Trim(RS.Fields("expdt").value & ""), "####-##-##")
            .Col = 8:
                      Dim strBldNm As String
                      
                      strBldNm = GetBuildNm(Trim("" & RS.Fields("centercd").value & ""))
                      If strBldNm <> "" Then
'                      If ObjBBSComCode.Building.Exists(Trim("" & Rs.Fields("centercd").value & "")) Then
'                         ObjBBSComCode.Building.KeyChange (Trim("" & Rs.Fields("centercd").value & ""))
                         .value = strBldNm 'ObjBBSComCode.Building.Fields("buildnm")
                      Else
                         .value = ObjSysInfo.BuildingNm
                      End If
                      If lblCenterNm.Caption = .value Then
                         .Col = -1: .Row = .MaxRows
                         .ForeColor = DCM_Gray
                      End If
            .Col = 9: .value = Trim(RS.Fields("compocd").value & "")
            
            If chkDup(.MaxRows, strBldno, RS.Fields("abo").value & "", RS.Fields("compocd").value & "") = False Then
                MsgBox " 중복된 혈액번호입니다.", vbCritical, Me.Caption
                .Row = .MaxRows
                .Col = -1
                .Action = 5
                .MaxRows = .MaxRows - 1
                Set RS = Nothing
                Set objSql = Nothing
                Exit Sub
            End If
        End With
        lblTotCnt.Caption = tblResult.DataRowCnt
    End If
    Set RS = Nothing
    Set objSql = Nothing
End Sub

Private Function chkDup(ByVal Prow As Long, ByVal pBldNo As String, ByVal pABO As String, ByVal pCompoCd As String) As Boolean
    Dim i As Long
    Dim strBldno As String
    Dim strABO As String
    Dim strCompocd As String

    If Prow = 1 Then chkDup = True: Exit Function

    With tblResult
        For i = 1 To .DataRowCnt - 1
            .Row = i
            .Col = 1: strBldno = Trim(.value)
            .Col = 2: strABO = Trim(.value)
            .Col = 9: strCompocd = Trim(.value)
            If Trim(pBldNo) = Trim(strBldno) And Trim(pABO) = Trim(strABO) And Trim(pCompoCd) = Trim(strCompocd) Then chkDup = False: Exit Function
        Next
    End With
    chkDup = True
End Function


Private Sub tblResult_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
        
    lngRow = Row
    tblResult.Row = lngRow
    tblResult.Col = -1
    tblResult.BackColor = &HC0C0C0
    lngRow = Row
    Set objPop = New clsPopupMenu
    With objPop
        .AddMenu MENU_DEL, "DELETE"
        .PopupMenus Me.hwnd
    End With
    Set objPop = Nothing
'    Set mnuPopup = frmControls.mnuPopup
'    Set mnuDelete = frmControls.mnuSub
'    mnuDelete.Caption = "Delete"
'    PopupMenu mnuPopup
'
'    Set mnuPopup = Nothing
'    Set mnuDelete = Nothing
    
    If lngRow Mod 2 = 0 Then
        tblResult.BackColor = &HE0E0E0
    Else
        tblResult.BackColor = -2147483643
    End If
End Sub

'Private Sub mnuDelete_Click()
'    With tblResult
'        .Row = lngRow
'        .Col = -1
'        .Action = 5
'        .MaxRows = .MaxRows - 1
'        lblTotCnt.Caption = .MaxRows
'    End With
'End Sub

Private Sub txtCenterCd_LostFocus()
    If txtCenterCd.Text = "" Then Exit Sub

    lblCenterNm.Caption = GetCenterNm(Trim(txtCenterCd.Text))
    If lblCenterNm.Caption = "" Then
        MsgBox "존재하지 않는 혈액센타코드입니다.", vbInformation, Me.Caption
        txtCenterCd.Text = ""
        txtCenterCd.SetFocus
        Exit Sub
    End If

    If txtCenterCd.Text = ObjSysInfo.BuildingCd Then
        MsgBox "같은 혈액센타를 선택 할 수 없습니다.", vbInformation, Me.Caption
        txtCenterCd.Text = ""
        lblCenterNm.Caption = ""
        txtCenterCd.SetFocus
    End If
End Sub

Private Sub Clear()
    txtCenterCd.Text = ""
    lblCenterNm.Caption = ""
    txtBldNo.Text = ""
    lblCompoCd.Caption = ""
    lblCompoNm.Caption = ""
    lblTotCnt.Caption = ""
    With tblResult
        .Row = -1
        .Col = -1
        .Text = ""
        .MaxRows = 50
    End With
    If txtCenterCd.Enabled Then txtCenterCd.SetFocus
    dtpFMonth.value = GetSystemDate
    dtpTMonth.value = GetSystemDate
    tbldata.MaxRows = 0
End Sub
