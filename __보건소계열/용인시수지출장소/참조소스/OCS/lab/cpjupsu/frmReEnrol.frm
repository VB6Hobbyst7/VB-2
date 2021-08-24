VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmReEnrol 
   BackColor       =   &H8000000A&
   Caption         =   "병동채혈 검체확인작업 화면"
   ClientHeight    =   7695
   ClientLeft      =   150
   ClientTop       =   945
   ClientWidth     =   11835
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7695
   ScaleWidth      =   11835
   WindowState     =   2  '최대화
   Begin Threed.SSPanel SSPanel7 
      Height          =   600
      Left            =   7380
      TabIndex        =   27
      Top             =   540
      Width           =   4200
      _Version        =   65536
      _ExtentX        =   7408
      _ExtentY        =   1058
      _StockProps     =   15
      Caption         =   " BarCode?"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Alignment       =   1
      Begin VB.TextBox txtBartext 
         BackColor       =   &H00C0E0FF&
         Height          =   330
         Left            =   945
         TabIndex        =   35
         Top             =   135
         Width           =   2085
      End
   End
   Begin Threed.SSPanel SSPanel6 
      Height          =   5640
      Left            =   90
      TabIndex        =   26
      Top             =   1170
      Width           =   2985
      _Version        =   65536
      _ExtentX        =   5265
      _ExtentY        =   9948
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Begin VB.OptionButton Option1 
         Caption         =   "정규Order"
         Height          =   270
         Left            =   1395
         TabIndex        =   29
         Top             =   900
         Width           =   1320
      End
      Begin VB.OptionButton Option2 
         Caption         =   "추가Order"
         Height          =   270
         Left            =   1395
         TabIndex        =   28
         Top             =   1215
         Value           =   -1  'True
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker dtBdate 
         Height          =   330
         Left            =   1260
         TabIndex        =   30
         Top             =   225
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24444931
         CurrentDate     =   36516
      End
      Begin MSForms.CommandButton cmdReset 
         Height          =   510
         Left            =   1260
         TabIndex        =   34
         Top             =   2745
         Width           =   1500
         Caption         =   "Clear"
         PicturePosition =   327683
         Size            =   "2646;900"
         Picture         =   "frmReEnrol.frx":0000
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdExec 
         Height          =   510
         Left            =   1260
         TabIndex        =   33
         Top             =   2250
         Width           =   1500
         Caption         =   "검체확인"
         PicturePosition =   327683
         Size            =   "2646;900"
         Picture         =   "frmReEnrol.frx":1792
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdQuery 
         Height          =   510
         Left            =   1260
         TabIndex        =   32
         Top             =   1755
         Width           =   1500
         Caption         =   "조회확인"
         PicturePosition =   327683
         Size            =   "2646;900"
         Picture         =   "frmReEnrol.frx":2F54
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin VB.Label Label9 
         Caption         =   "접수일자"
         Height          =   195
         Left            =   270
         TabIndex        =   31
         Top             =   270
         Width           =   825
      End
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   285
      Left            =   8010
      TabIndex        =   1
      Top             =   135
      Visible         =   0   'False
      Width           =   3525
      _Version        =   65536
      _ExtentX        =   6218
      _ExtentY        =   503
      _StockProps     =   15
      Caption         =   "검체번호 개별 도착 확인"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Alignment       =   0
      Begin Threed.SSPanel SSPanel2 
         Height          =   2625
         Left            =   495
         TabIndex        =   2
         Top             =   1080
         Width           =   2715
         _Version        =   65536
         _ExtentX        =   4789
         _ExtentY        =   4630
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   2
         Alignment       =   0
         Begin VB.TextBox txtPtno 
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   855
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Text            =   "txtPtno"
            Top             =   225
            Width           =   1140
         End
         Begin VB.TextBox txtDrname 
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   855
            Locked          =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            Text            =   "txtDrname"
            Top             =   2205
            Width           =   1140
         End
         Begin VB.TextBox txtRoom 
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   855
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Text            =   "txtRoom"
            Top             =   885
            Width           =   1140
         End
         Begin VB.TextBox txtDeptName 
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   855
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            Text            =   "txtDeptname"
            Top             =   1875
            Width           =   1140
         End
         Begin VB.TextBox txtBirthDay 
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   1350
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            Text            =   "txtBirthDay"
            Top             =   1545
            Width           =   1140
         End
         Begin VB.TextBox txtAgeYY 
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   855
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            Text            =   "txtAgeYY"
            Top             =   1545
            Width           =   465
         End
         Begin VB.TextBox txtSex 
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   855
            Locked          =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            Text            =   "txtSex"
            Top             =   1215
            Width           =   465
         End
         Begin VB.TextBox txtSname 
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   855
            Locked          =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            Text            =   "txtSname"
            Top             =   555
            Width           =   1140
         End
         Begin VB.Label Label2 
            Caption         =   "등록번호"
            Height          =   240
            Left            =   90
            TabIndex        =   17
            Top             =   270
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "환자명"
            Height          =   240
            Left            =   90
            TabIndex        =   16
            Top             =   585
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "병실"
            Height          =   240
            Left            =   90
            TabIndex        =   15
            Top             =   900
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "성별"
            Height          =   195
            Left            =   90
            TabIndex        =   14
            Top             =   1260
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "나이"
            Height          =   240
            Left            =   90
            TabIndex        =   13
            Top             =   1575
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "과"
            Height          =   240
            Left            =   90
            TabIndex        =   12
            Top             =   1935
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "의사"
            Height          =   240
            Left            =   90
            TabIndex        =   11
            Top             =   2250
            Width           =   735
         End
      End
      Begin FPSpreadADO.fpSpread sprConfirm 
         Height          =   4200
         Left            =   3240
         TabIndex        =   18
         Top             =   1080
         Width           =   7080
         _Version        =   196608
         _ExtentX        =   12488
         _ExtentY        =   7408
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   8
         MaxRows         =   50
         ScrollBars      =   2
         SpreadDesigner  =   "frmReEnrol.frx":3836
         Appearance      =   2
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   555
         Left            =   495
         TabIndex        =   19
         Top             =   495
         Width           =   4695
         _Version        =   65536
         _ExtentX        =   8281
         _ExtentY        =   979
         _StockProps     =   15
         BackColor       =   8421376
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         Begin VB.TextBox txtBarCode 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1215
            TabIndex        =   20
            Top             =   135
            Width           =   2175
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808000&
            Caption         =   "BarCode :"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   135
            TabIndex        =   21
            Top             =   180
            Width           =   915
         End
      End
      Begin MSForms.CommandButton cmdOk 
         Height          =   555
         Left            =   5310
         TabIndex        =   23
         Top             =   360
         Width           =   1905
         Caption         =   " 검체확인[F4]"
         PicturePosition =   327683
         Size            =   "3360;979"
         Picture         =   "frmReEnrol.frx":41BE
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdClear 
         Height          =   555
         Left            =   7200
         TabIndex        =   22
         Top             =   360
         Width           =   1905
         Caption         =   "   Clear[F1]"
         PicturePosition =   327683
         Size            =   "3360;979"
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
   End
   Begin FPSpreadADO.fpSpread sprOrder 
      Height          =   6045
      Left            =   3195
      TabIndex        =   24
      Top             =   1170
      Width           =   8430
      _Version        =   196608
      _ExtentX        =   14870
      _ExtentY        =   10663
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   20
      ScrollBars      =   2
      SpreadDesigner  =   "frmReEnrol.frx":5980
      UserResize      =   1
      Appearance      =   2
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   555
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   3525
      _Version        =   65536
      _ExtentX        =   6218
      _ExtentY        =   979
      _StockProps     =   15
      Caption         =   "병동,ER 검체 확인"
      ForeColor       =   65535
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "궁서체"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSForms.CommandButton cmdSelect 
      Height          =   375
      Left            =   3600
      TabIndex        =   25
      Top             =   765
      Width           =   1365
      Caption         =   "▼전체선택"
      Size            =   "2408;661"
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
   Begin VB.Menu mnuLabPass 
      Caption         =   "개별확인"
   End
End
Attribute VB_Name = "frmReEnrol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    
    For i = 0 To Me.Count - 1
        If TypeOf Me.Controls(i) Is VB.TextBox Then
            Me.Controls(i).Text = ""
        End If
    Next
    
    sprConfirm.Row = 1
    sprConfirm.Row2 = sprConfirm.DataRowCnt
    sprConfirm.Col = 1
    sprConfirm.Col2 = sprConfirm.MaxCols
    sprConfirm.BlockMode = True
    sprConfirm.Action = ActionClearText
    sprConfirm.BlockMode = False
    
    

    
    
End Sub

Private Sub cmdExec_Click()
    Dim sOrderRowID     As String
    Dim sPtno           As String
    Dim sCollDate       As String
    Dim sCollHH         As String
    Dim sCollMM         As String
    Dim sOrderno        As String
    Dim iCount          As String
    Dim sDept           As String
    Dim iMatchno        As Integer
    
        
    If sprOrder.DataRowCnt = 0 Then Exit Sub
    
    iCount = 0
    For i = 1 To sprOrder.DataRowCnt
        sprOrder.Row = i
        sprOrder.Col = 1
        If sprOrder.Value = True Then
            sprOrder.Col = 19:  sPtno = sprOrder.Text
            sprOrder.Col = 8:  sCollDate = sprOrder.Text
            sprOrder.Col = 9:  sCollHH = sprOrder.Text
            sprOrder.Col = 10: sCollMM = sprOrder.Text
            sprOrder.Col = 11: sOrderRowID = sprOrder.Text
            sprOrder.Col = 12: sOrderno = sprOrder.Text
            sprOrder.Col = 20: iMatchno = Val(sprOrder.Text)
            GoSub General_Data_Update
            GoSub Order_Data_Update
            iCount = iCount + 1
        End If
    Next
    
    MsgBox "모든 작업이 끝났습니다!........" & vbCrLf & _
           "총 " & sprOrder.DataRowCnt & " 개의 Data를 준비하여 " & vbCrLf & _
           iCount & " 개의 Order를 검체확인 하였습니다", vbInformation
    Exit Sub
    


General_Data_Update:
    Dim sCurrentDate        As String
    Dim sCurrDate           As String
    Dim sCurrHH             As String
    Dim sCurrMM             As String
    
    
    sCurrentDate = Dual_Date_Get("yyyy-MM-dd hh24:mi")
    sCurrDate = Left(sCurrentDate, 10)
    sCurrHH = Mid(sCurrentDate, 12, 2)
    sCurrMM = Right(sCurrentDate, 2)
    
    strSql = ""
    strSql = strSql & " UPDATE TWEXAM_General"
    strSql = strSql & " SET    GBCH     = 'Y',"
    strSql = strSql & "        GeomsaDt = TO_DATE('" & sCurrDate & "','yyyy-MM-dd'),"
    strSql = strSql & "        GeomsaT1 = '" & sCurrHH & "',"
    strSql = strSql & "        GeomsaT2 = '" & sCurrMM & "'"
    strSql = strSql & " WHERE  JeobsuDt = TO_DATE('" & sCollDate & "','yyyy-MM-dd')"
    strSql = strSql & " AND    JeobsuT1 = '" & sCollHH & "'"
    strSql = strSql & " AND    JeobsuT2 = '" & sCollMM & "'"
    strSql = strSql & " AND    Matchno  =  " & iMatchno
    strSql = strSql & " AND    Ptno     = '" & sPtno & "'"
    
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    Return



Order_Data_Update:
    strSql = ""
    strSql = strSql & " UPDATE TW_MIS_EXAM.TWEXAM_Order"
    strSql = strSql & " SET    GBCH = 'Y',"
    strSql = strSql & "        COLLid = " & Val(GstrIdnumber)
    strSql = strSql & " WHERE  ROWID = '" & sOrderRowID & "'"
    
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    Return
    
    
End Sub

Private Sub cmdOk_Click()
    Dim sRowID      As String
    Dim nOrderno    As String
    Dim sOrderno    As String
    
    
    If sprConfirm.DataRowCnt = 0 Then
        MsgBox "해당 접수 Data 가 하나도 없습니다!.. 접수한 Data를 먼저확인하세요"
        Exit Sub
    End If
    
    For i = 1 To sprConfirm.DataRowCnt
        sprConfirm.Row = i
        sprConfirm.Col = 1: sRowID = sprConfirm.Text
        sprConfirm.Col = 8: sOrderno = sprConfirm.Text
        GoSub Update_General_GbCH
    Next
    
    Call cmdClear_Click
    txtBarCode.SetFocus
    
    Exit Sub
    

Update_General_GbCH:
    strSql = ""
    strSql = strSql & " UPDATE TWEXAM_General"
    strSql = strSql & " SET    GbCH  = 'Y'"
    strSql = strSql & " WHERE  RowID = '" & sRowID & "'"
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    
  'Exam_Order Update
    Dim sJDt          As String
    Dim sSLno1        As String
    Dim sToDate       As String
    Dim nToHH         As Integer
    Dim nToMM         As Integer
    
    
    'sJDt = Left(txtBarCode.Text, 8)
    
    sJDt = convLabnoToExpand(Left(txtBarCode.Text, 5))
    
    sSLno1 = Val(Mid(txtBarCode.Text, 6, 2))

    sToDate = Dual_Date_Get("yyyy-MM-dd")
    nToHH = Val(Dual_Date_Get("hh"))
    nToMM = Val(Dual_Date_Get("mi"))
    
    strSql = ""
    strSql = strSql & " UPDATE TW_MIS_EXAM.TWEXAM_Order"
    strSql = strSql & " SET    GbCH      =  'Y',"
    strSql = strSql & "        CoLLDate  =   TO_DATE('" & sToDate & "','YYYY-MM-DD'),"
    strSql = strSql & "        CoLLHH    =   " & nToHH & ","
    strSql = strSql & "        CoLLMM    =   " & nToMM & ","
    strSql = strSql & "        CoLLid    =   " & Val(GstrIdnumber)
    strSql = strSql & " WHERE  JeobsuDt  =   TO_DATE('" & sJDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    SLipno1   =  '" & Val(sSLno1) & "'"
    strSql = strSql & " AND    Orderno   =   " & Val(sOrderno) & ""
    
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    Return
    
    
    
End Sub

Private Sub cmdQuery_Click()
    Dim sBDate      As String
    
    Dim sPtno       As String
    Dim sCollDate   As String
    Dim sCollHH     As String
    Dim sCollMM     As String
    Dim sOrderRowID As String
    Dim sOrderno    As String
    Dim sSLipno1    As String
    Dim sTmpTEXT    As String
    
    
    Call Spread_Set_Clear(sprOrder)
    cmdSelect.Caption = "▼전체선택"
    
    
    sBDate = Format(dtBdate.Value, "yyyy-MM-dd")
    
    strSql = ""
    strSql = strSql & " SELECT DISTINCT b.RoutinNM ItemName, a.*, a.RowID OrderRowID, "
    strSql = strSql & "        TO_CHAR(a.CollDate,'yyyy-MM-dd') COLLDate, c.Sname,"
    strSql = strSql & "        a.Ptno PatientNO, a.SLipno1 SLno1, a.Matchno"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Order   a,"
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Routine b,"
    strSql = strSql & "        TWEXAM_IDNOMST c "
    strSql = strSql & " WHERE  a.CollDate = TO_DATE('" & sBDate & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.GBCH    IN ('1','2')"
    strSql = strSql & " AND    a.JeobsuYn = '*'"
    
    If Option1.Value = True Then strSql = strSql & "  AND  a.OrderGb IN ('X','Y','Z',' ')" '정규Order
    If Option2.Value = True Then strSql = strSql & " AND  a.OrderGb IN ('F','T','M','A')" '추가Order
    
    strSql = strSql & " AND    a.SLipno1  > 0 "
'C    strSql = strSql & " AND    a.SLipno1  < 52"
    strSql = strSql & " AND    a.SLipno1  < 90"
    'strSql = strSql & " AND    a.GBIO     = 'I'"         '고놈에 응급실 때문에 일단 막았음
    strSql = strSql & " AND    a.ItemCd   = b.RoutinCd"
    strSql = strSql & " AND    a.Ptno     = c.Ptno(+)"
    
    strSql = strSql & " UNION ALL"
    strSql = strSql & " SELECT b.ItemNM ItemName, a.*, a.RowID OrderRowID, "
    strSql = strSql & "        TO_CHAR(a.CollDate,'yyyy-MM-dd') COLLDate, c.Sname,"
    strSql = strSql & "        a.Ptno PatientNO, a.SLipno1 SLno1, a.Matchno"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Order   a,"
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_itemML  b,"
    strSql = strSql & "        TWEXAM_IDNOMST c"
    strSql = strSql & " WHERE  a.CollDate = TO_DATE('" & sBDate & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.GbCh    IN ('1','2')"
    strSql = strSql & " AND    a.JeobsuYn = '*'"
    
    If Option1.Value = True Then strSql = strSql & "  AND  a.OrderGb IN ('X','Y','Z',' ')" '정규Order
    If Option2.Value = True Then strSql = strSql & " AND  a.OrderGb IN ('F','T','M','A')" '추가Order
    
    strSql = strSql & " AND    a.SLipno1  > 0 "
'C    strSql = strSql & " AND    a.SLipno1  < 52"
    strSql = strSql & " AND    a.SLipno1  < 90 "
    'strSql = strSql & " AND    a.GBIO     = 'I'"
    strSql = strSql & " AND    a.ItemCd   = b.Codeky"
    strSql = strSql & " AND    a.Ptno     = c.Ptno(+)"
    strSql = strSql & " Order by    PatientNO, SLno1"
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        sprOrder.Row = sprOrder.DataRowCnt + 1
        sprOrder.Col = 1:  sprOrder.Value = False
        If sTmpTEXT <> adoSet.Fields("Patientno").Value & "" Then
            sprOrder.Col = 2:  sprOrder.Text = adoSet.Fields("Patientno").Value & ""
                               sPtno = adoSet.Fields("Patientno").Value & ""
            sprOrder.Col = 3:  sprOrder.Text = adoSet.Fields("Sname").Value & ""
            sprOrder.Col = 4:  sprOrder.Text = adoSet.Fields("Sex").Value & ""
            sprOrder.Col = 5:  sprOrder.Text = adoSet.Fields("AgeYY").Value & ""
            sprOrder.Col = 6:  sprOrder.Text = adoSet.Fields("RoomCode").Value & ""
        End If
        
        sPtno = adoSet.Fields("Patientno").Value & ""
        sprOrder.Col = 7:  sprOrder.Text = adoSet.Fields("ItemName").Value & ""
        sprOrder.Col = 8:  sprOrder.Text = adoSet.Fields("CoLLDate").Value & ""
                           sCollDate = sprOrder.Text
        sprOrder.Col = 9:  sprOrder.Text = adoSet.Fields("CoLLHH").Value & ""
                           sCollHH = sprOrder.Text
        sprOrder.Col = 10: sprOrder.Text = adoSet.Fields("CoLLMM").Value & ""
                           sCollMM = sprOrder.Text
        sprOrder.Col = 11: sprOrder.Text = adoSet.Fields("OrderRowID").Value & ""
        sprOrder.Col = 12: sprOrder.Text = adoSet.Fields("Orderno").Value & ""
                           sOrderno = sprOrder.Text
                           
        sprOrder.Col = 13
        Select Case Trim(adoSet.Fields("GbCH").Value & "")
            Case "1": sprOrder.Text = "추가Order"
            Case "2": sprOrder.Text = "정규Order"
        End Select
        
        sSLipno1 = adoSet.Fields("SLno1").Value & ""
        sprOrder.Col = 18: sprOrder.Text = adoSet.Fields("DeptCode").Value & ""
        sprOrder.Col = 19: sprOrder.Text = adoSet.Fields("Patientno").Value & ""
        sprOrder.Col = 20: sprOrder.Text = adoSet.Fields("Matchno").Value & ""
        GoSub General_Labno_Select
        
        sTmpTEXT = adoSet.Fields("Patientno").Value & ""
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    txtBartext.Text = ""
    txtBartext.SetFocus
    
    
    
    Exit Sub
    
    

General_Labno_Select:
    Dim adoGn       As ADODB.Recordset
    Dim sSumKey     As String
    
    strSql = ""
    strSql = strSql & " SELECT TO_CHAR(JeobsuDt, 'yyyy-MM-dd') JeobsuDt,"
    strSql = strSql & "        SLipno1, SLipno2"
    strSql = strSql & " FROM   TWEXAM_General"
    strSql = strSql & " WHERE  JeobsuDt = TO_DATE('" & sCollDate & "','yyyy-MM-dd')"
    strSql = strSql & " AND    JeobsuT1 = " & Val(sCollHH)
    strSql = strSql & " AND    JeobsuT2 = " & Val(sCollMM)
    strSql = strSql & " AND    SLipno1  = " & Val(sSLipno1)
    strSql = strSql & " AND    Ptno     = '" & sPtno & "'"
    
    If False = adoSetOpen(strSql, adoGn) Then Return
    
    sSumKey = ""
    sprOrder.Row = sprOrder.Row
    sprOrder.Col = 14: sprOrder.Text = adoGn.Fields("JeobsuDt").Value & ""
                       sSumKey = convLabnoToComp(Replace(sprOrder.Text, "-", ""))
    sprOrder.Col = 15: sprOrder.Text = adoGn.Fields("SLipno1").Value & ""
                       sSumKey = sSumKey & Trim(sprOrder.Text)
    sprOrder.Col = 16: sprOrder.Text = adoGn.Fields("SLipno2").Value & ""
                       sSumKey = sSumKey & Format(Trim(sprOrder.Text), "00000")
    
    sprOrder.Col = 17: sprOrder.Text = sSumKey
    
    
    Return
    
End Sub

Private Sub cmdReset_Click()
    
    dtBdate.Value = Dual_Date_Get("yyyy-MM-dd")
    Call Spread_Set_Clear(sprOrder)
    cmdSelect.Caption = "▼전체선택"
    
End Sub

Private Sub cmdSelect_Click()
        
    
    If cmdSelect.Caption = "▼전체선택" Then
    
        For i = 1 To sprOrder.DataRowCnt
            sprOrder.Row = i
            sprOrder.Col = 1
            sprOrder.Value = True
        Next
        cmdSelect.Caption = "▼전체해제"
    Else
        For i = 1 To sprOrder.DataRowCnt
            sprOrder.Row = i
            sprOrder.Col = 1
            sprOrder.Value = False
        Next
        cmdSelect.Caption = "▼전체선택"
    End If
    
    
    
End Sub

Private Sub Form_Activate()
    Me.WindowState = vbMaximized
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    
    Select Case KeyCode
        Case vbKeyF1: Call cmdClear_Click
        Case vbKeyF4: Call cmdOk_Click

    End Select
    
End Sub

Private Sub Form_Load()
        
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
        
    
    Call cmdReset_Click
    
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

Private Sub mnuLabPass_Click()
    
    frmReJupsu.Show vbModal
    
    
End Sub

Private Sub sprOrder_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    
    If Row = 0 Then
        GoSub Set_Sort_Spread_sprOrder
    End If
    Exit Sub
    
Set_Sort_Spread_sprOrder:
    sprOrder.Col = 1
    sprOrder.Col2 = sprOrder.MaxCols
    sprOrder.Row = 1
    sprOrder.Row2 = sprOrder.DataRowCnt
    
    sprOrder.SortBy = SS_SORT_BY_ROW
    sprOrder.SortKey(1) = 1
    
    If sprOrder.SortKeyOrder(1) = SortKeyOrderAscending Then
        sprOrder.SortKeyOrder(1) = SortKeyOrderDescending
    Else
        sprOrder.SortKeyOrder(1) = SortKeyOrderAscending
    End If
    
    sprOrder.Action = SS_ACTION_SORT
    Return
    
    
End Sub

Private Sub txtBarCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sJeobsuDt        As String
    Dim iSLipno1         As Integer
    Dim iSLipno2         As Integer
    
    
    If KeyCode = vbKeyReturn Then
        If Trim(txtBarCode.Text) = "" Then Exit Sub
        txtBarCode.Tag = txtBarCode.Text
        Call cmdClear_Click
        txtBarCode.Text = txtBarCode.Tag
        
        Select Case Len(Trim(txtBarCode.Text))
            Case 12
                sJeobsuDt = convLabnoToExpand(Left(txtBarCode.Text, 5))
                iSLipno1 = Val(Mid(txtBarCode.Text, 6, 2))
                iSLipno2 = Val(Mid(txtBarCode.Text, 8, 5))
            Case 15
                sJeobsuDt = Left(txtBarCode.Text, 8)
                iSLipno1 = Val(Mid(txtBarCode.Text, 9, 2))
                iSLipno2 = Val(Mid(txtBarCode.Text, 11, 5))
        End Select
        
        GoSub Get_General_Data
        If Trim(txtPtno.Text) <> "" Then
            Call txtPtno_KeyDown(vbKeyReturn, 1)
        End If
        cmdOk.SetFocus
        
    End If
    
    Exit Sub
    

Get_General_Data:
    strSql = ""
    strSql = strSql & " SELECT a.RowID RwID, a.*,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDt,'YYYY-MM-DD') JeobsuDt,"
    strSql = strSql & "        b.Codenm SLipName, c.Codenm SampleName, a.Orderno"
    strSql = strSql & " FROM   TWEXAM_General a,"
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Specode b,"
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Sample  c "
    strSql = strSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyyMMdd')"
    strSql = strSql & " AND    a.SLipno1  = " & iSLipno1
    strSql = strSql & " AND    a.SLipno2  = " & iSLipno2
    strSql = strSql & " AND    a.GBCH    IN  ('1','2')"  '1=병동에서 접수, 2=정규채혈
    strSql = strSql & " AND    a.SLipno1  =  TO_Number(b.Codeky)"
    strSql = strSql & " AND    b.Codegu   =   '12'"
    strSql = strSql & " AND    a.GeomchCd =  c.Code(+)"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sprConfirm.Row = sprConfirm.DataRowCnt + 1
        sprConfirm.Col = 1: sprConfirm.Text = adoSet.Fields("RwID").Value & ""
        sprConfirm.Col = 2: sprConfirm.Text = adoSet.Fields("JeobsuDt").Value & ""
        sprConfirm.Col = 3: sprConfirm.Text = adoSet.Fields("SampleName").Value & ""
        sprConfirm.Col = 4: sprConfirm.Text = adoSet.Fields("SLipName").Value & ""
        sprConfirm.Col = 5: sprConfirm.Text = adoSet.Fields("SLipno2").Value & ""
        sprConfirm.Col = 6: sprConfirm.Text = adoSet.Fields("GbCh").Value & ""
        
        Select Case adoSet.Fields("GbCh").Value & ""
            Case "1": sprConfirm.Col = 7: sprConfirm.Text = "병동 Or ER 접수"
            Case "2": sprConfirm.Col = 7: sprConfirm.Text = "정규채혈"
        End Select
        
        txtPtno.Text = adoSet.Fields("Ptno").Value & ""
        sprConfirm.Col = 8: sprConfirm.Text = adoSet.Fields("Orderno").Value & ""
        adoSet.MoveNext
        
    Loop
    Call adoSetClose(adoSet)
    
        
    Return

End Sub

Private Sub txtBartext_GotFocus()
    
    txtBartext.SelStart = 0
    txtBartext.SelLength = Len(txtBartext.Text)
    
End Sub

Private Sub txtBartext_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sKey        As String
    
    
    'if True then False AND  False Then True
    
    If KeyCode = vbKeyReturn Then
        
        For i = 1 To Me.sprOrder.DataRowCnt
            sprOrder.Row = i
            sprOrder.Col = 17: sKey = sprOrder.Text
            If Trim(sprOrder.Text) = (txtBartext.Text) Then
                sprOrder.Col = 1
                If sprOrder.Value = True Then
                    sprOrder.Value = False
                Else
                    sprOrder.Value = True
                End If
                sprOrder.Action = ActionGotoCell
                sprOrder.Col = 1
                sprOrder.Action = ActionActiveCell
                
            End If
        Next
        txtBartext.Text = ""
        txtBartext.SetFocus
    
    End If
    
    
End Sub

Private Sub txtPtno_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        GoSub Check_Ptno_Text
        GoSub Main_Search_Data
    End If
    Exit Sub
    


Check_Ptno_Text:
    
    If Trim(txtPtno.Text) = "" Then Exit Sub
    If Not IsNumeric(txtPtno.Text) Then Exit Sub
    txtPtno.Text = Format(txtPtno.Text, "00000000")
        
    Return
    
    
Main_Search_Data:
    strSql = ""
    strSql = strSql & " SELECT b.WardCode, a.*, c.DeptnameK, d.Drname"
    strSql = strSql & " FROM   TWEXAM_IDNOMST a,"
    strSql = strSql & "        TW_MIS_PMPA.TWBas_Room     b,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT     c,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR   d "
    strSql = strSql & " WHERE  a.Ptno     = '" & txtPtno.Text & "'"
    strSql = strSql & " AND    a.GbIO     = 'I'"
    strSql = strSql & " AND    a.RoomCode = b.RoomCode"
    strSql = strSql & " AND    a.DeptCode = c.DeptCode(+)"
    strSql = strSql & " AND    a.DrCode   = d.DrCode(+)"
    If False = adoSetOpen(strSql, adoSet) Then
        MsgBox "등록번호 " & txtPtno.Text & " 는(은) 접수된 Data 가 없습니다!..."
        Call cmdClear_Click
        Exit Sub
    Else
        txtSname.Text = adoSet.Fields("Sname").Value & ""
        txtRoom.Text = adoSet.Fields("RoomCode").Value & ""
        txtSex.Text = adoSet.Fields("Sex").Value & ""
        txtAgeYY.Text = adoSet.Fields("AgeYY").Value & ""
        txtBirthDay.Text = adoSet.Fields("BirthDay").Value & ""
        txtDeptName.Text = adoSet.Fields("DeptnameK").Value & ""
        txtDrname.Text = adoSet.Fields("Drname").Value & ""
        Call adoSetClose(adoSet)
    End If
        
    Return

End Sub

