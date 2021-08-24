VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#8.0#0"; "FPSPRU80.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form frm업체기초자료 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "업체기초자료"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14535
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   14535
   ShowInTaskbar   =   0   'False
   Begin SSSplitter.SSSplitter splMain 
      Height          =   7545
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   13309
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   7
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   1
      Locked          =   -1  'True
      PaneTree        =   "frm업체기초자료.frx":0000
      Begin FPUSpreadADO.fpSpread spList 
         Height          =   6090
         Left            =   6240
         TabIndex        =   46
         Top             =   750
         Width           =   8265
         _Version        =   524288
         _ExtentX        =   14579
         _ExtentY        =   10742
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   -2147483633
         MaxCols         =   5
         OperationMode   =   3
         SpreadDesigner  =   "frm업체기초자료.frx":00D2
         UserResize      =   0
         AppearanceStyle =   0
      End
      Begin Threed.SSPanel SSPanel26 
         Height          =   570
         Left            =   6240
         TabIndex        =   45
         Top             =   6945
         Width           =   8265
         _ExtentX        =   14579
         _ExtentY        =   1005
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   1380
            TabIndex        =   48
            Text            =   "Text1"
            Top             =   150
            Width           =   5385
         End
         Begin Threed.SSPanel SSPanel27 
            Height          =   300
            Left            =   150
            TabIndex        =   47
            Top             =   150
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "검색어"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSCommand cmdFInd 
            Height          =   420
            Left            =   6930
            TabIndex        =   49
            Top             =   90
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   741
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "검색"
            ButtonStyle     =   2
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   6090
         Left            =   30
         TabIndex        =   24
         Top             =   750
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   10742
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkDelete 
            Caption         =   "삭제"
            Height          =   300
            Left            =   1560
            TabIndex        =   14
            Top             =   4860
            Width           =   975
         End
         Begin VB.TextBox txtHp 
            Height          =   300
            Left            =   4290
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   3000
            Width           =   1575
         End
         Begin VB.TextBox txtAddr2 
            Height          =   300
            Left            =   1410
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   2340
            Width           =   4455
         End
         Begin VB.TextBox txtRemark 
            Height          =   630
            Left            =   1410
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   15
            Text            =   "frm업체기초자료.frx":072B
            Top             =   5280
            Width           =   4455
         End
         Begin VB.TextBox txtPost 
            Height          =   300
            Left            =   1410
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   2010
            Width           =   975
         End
         Begin VB.TextBox txtBankNo 
            Height          =   300
            Left            =   1410
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   3750
            Width           =   4455
         End
         Begin VB.TextBox txtBank 
            Height          =   300
            Left            =   1410
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   3420
            Width           =   4455
         End
         Begin VB.TextBox txtMan 
            Height          =   300
            Left            =   1410
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   3000
            Width           =   1575
         End
         Begin VB.TextBox txtFax 
            Height          =   300
            Left            =   4290
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   2670
            Width           =   1575
         End
         Begin VB.TextBox txtTel 
            Height          =   300
            Left            =   1410
            TabIndex        =   8
            Text            =   "123456789012345"
            Top             =   2670
            Width           =   1575
         End
         Begin VB.TextBox txtAddr1 
            Height          =   300
            Left            =   2400
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   2010
            Width           =   3465
         End
         Begin VB.TextBox txtItem 
            Height          =   300
            Left            =   1410
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   1590
            Width           =   4455
         End
         Begin VB.TextBox txtType 
            Height          =   300
            Left            =   1410
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   1260
            Width           =   4455
         End
         Begin VB.TextBox txtId 
            Height          =   300
            Left            =   4290
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   930
            Width           =   1575
         End
         Begin VB.TextBox txtMng 
            Height          =   300
            Left            =   1410
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   930
            Width           =   1575
         End
         Begin VB.TextBox txtNm 
            Height          =   300
            Left            =   1410
            TabIndex        =   0
            Text            =   "Text1"
            Top             =   510
            Width           =   4455
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   300
            Left            =   180
            TabIndex        =   25
            Top             =   180
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "업체번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel6 
            Height          =   300
            Left            =   180
            TabIndex        =   26
            Top             =   510
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "업 체 명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   300
            Left            =   180
            TabIndex        =   27
            Top             =   930
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "대표자명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   300
            Left            =   3060
            TabIndex        =   28
            Top             =   930
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "사업자번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel9 
            Height          =   300
            Left            =   180
            TabIndex        =   29
            Top             =   1260
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "업 태"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel10 
            Height          =   300
            Left            =   180
            TabIndex        =   30
            Top             =   1590
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "종 목"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel11 
            Height          =   630
            Left            =   180
            TabIndex        =   31
            Top             =   2010
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1111
            _Version        =   262144
            Caption         =   "주 소"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel12 
            Height          =   300
            Left            =   180
            TabIndex        =   32
            Top             =   2670
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "전화번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel13 
            Height          =   300
            Left            =   3060
            TabIndex        =   33
            Top             =   2670
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "팩스번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel14 
            Height          =   300
            Left            =   180
            TabIndex        =   34
            Top             =   3000
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "담당자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel15 
            Height          =   300
            Left            =   180
            TabIndex        =   35
            Top             =   3420
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "결재은행"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel16 
            Height          =   300
            Left            =   180
            TabIndex        =   36
            Top             =   3750
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "계좌번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel17 
            Height          =   300
            Left            =   180
            TabIndex        =   37
            Top             =   4200
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "등록일시"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel18 
            Height          =   300
            Left            =   180
            TabIndex        =   38
            Top             =   4530
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "수정일시"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel19 
            Height          =   300
            Left            =   180
            TabIndex        =   39
            Top             =   4860
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "사용안함"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel20 
            Height          =   630
            Left            =   180
            TabIndex        =   40
            Top             =   5280
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1111
            _Version        =   262144
            Caption         =   "비고사항"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel lblNo 
            Height          =   300
            Left            =   1410
            TabIndex        =   41
            Top             =   180
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "업체번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel22 
            Height          =   300
            Left            =   3060
            TabIndex        =   42
            Top             =   3000
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "핸드폰번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel lblWrtdt 
            Height          =   300
            Left            =   1410
            TabIndex        =   43
            Top             =   4200
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "2012-06-15 15:30:15"
            BevelOuter      =   1
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel lblModdt 
            Height          =   300
            Left            =   1410
            TabIndex        =   44
            Top             =   4530
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "2012-06-15 15:30:15"
            BevelOuter      =   1
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   615
         Left            =   30
         TabIndex        =   17
         Top             =   30
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   1085
         _Version        =   262144
         Font3D          =   5
         ForeColor       =   65535
         BackColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " ▒ 업체기초정보"
         BevelOuter      =   1
         BevelInner      =   2
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   570
         Left            =   30
         TabIndex        =   18
         Top             =   6945
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   1005
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSCommand cmdSave 
            Height          =   420
            Left            =   2550
            TabIndex        =   19
            Top             =   90
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   741
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "저장(&S)"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand cmdDelete 
            Height          =   420
            Left            =   3660
            TabIndex        =   20
            Top             =   90
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   741
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "삭제(&D)"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand cmdClear 
            Height          =   420
            Left            =   1440
            TabIndex        =   21
            Top             =   90
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   741
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "화면지움"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand cmdClose 
            Height          =   420
            Left            =   4770
            TabIndex        =   22
            Top             =   90
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   741
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "닫기(&X)"
            ButtonStyle     =   2
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   615
         Left            =   6240
         TabIndex        =   23
         Top             =   30
         Width           =   8265
         _ExtentX        =   14579
         _ExtentY        =   1085
         _Version        =   262144
         Font3D          =   5
         ForeColor       =   65535
         BackColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " ▒ 업체등록현황"
         BevelOuter      =   1
         BevelInner      =   2
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
End
Attribute VB_Name = "frm업체기초자료"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cCst As clsMstCust

Private Sub psFieldLengthDefine()

    gSql = "select * from mstCUST where custcd = 0"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            txtNm.MaxLength = .Fields("custnm").DefinedSize
            txtId.MaxLength = .Fields("custid").DefinedSize
            txtMng.MaxLength = .Fields("custmng").DefinedSize
            txtType.MaxLength = .Fields("custtype").DefinedSize
            txtItem.MaxLength = .Fields("custitem").DefinedSize
            txtPost.MaxLength = .Fields("postno").DefinedSize
            txtAddr1.MaxLength = .Fields("addr1").DefinedSize
            txtAddr2.MaxLength = .Fields("addr2").DefinedSize
            txtMan.MaxLength = .Fields("custman").DefinedSize
            txtBank.MaxLength = .Fields("banknm").DefinedSize
            txtBankNo.MaxLength = .Fields("bankno").DefinedSize
            txtRemark.MaxLength = .Fields("remark").DefinedSize
            .Close
        End If
    End With

End Sub

Private Sub psListRefresh()
Dim sRow As Long

    gSql = "select custcd, custnm, custid, telno, custman from mstCUST where custcd > 0"
    If Len(txtFind.Text) > 0 Then
        gSql = gSql & " and custnm like '%" & Trim(txtFind.Text) & "%'"
    End If
    gSql = gSql & " order by custnm"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                Call gsSpreadClear(spList, .RecordCount, True)
                While (Not .EOF)
                    sRow = sRow + 1
                    
                    spList.SetText 1, sRow, "" & .Fields("custcd").Value
                    spList.SetText 2, sRow, .Fields("custnm").Value
                    spList.SetText 3, sRow, .Fields("custid").Value
                    spList.SetText 4, sRow, .Fields("telno").Value
                    spList.SetText 5, sRow, .Fields("custman").Value
                    
                    .MoveNext
                Wend
            Else
                Call gsSpreadClear(spList, 1, True)
            End If
            .Close
        End If
    End With

End Sub

Private Sub cmdClear_Click()

    Call gsFieldClear(Me)
    
    cmdSave.Enabled = False
    cmdDelete.Enabled = False
    
    txtNm.SetFocus
    
End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdDelete_Click()

    MousePointer = vbHourglass
    If Val(lblNo.Caption) > 0 Then
        If cCst.cfDelete(Val(lblNo.Caption)) Then
            Call cmdClear_Click
            Call psListRefresh
        End If
    End If
    MousePointer = vbDefault
    
End Sub

Private Sub cmdFInd_Click()

    MousePointer = vbHourglass
    Call psListRefresh
    MousePointer = vbDefault

End Sub

Private Sub cmdSave_Click()

    On Error GoTo ErrcmdSave
    If Len(txtNm.Text) = 0 Then
        MsgBox "업체명을 입력하세요. !", vbCritical
        txtNm.SetFocus
        Exit Sub
    End If
    
    MousePointer = vbHourglass
    With cCst
        .custcd = Val(lblNo.Caption)
        .addr1 = Trim(txtAddr1.Text)
        .addr2 = Trim(txtAddr2.Text)
        .banknm = Trim(txtBank.Text)
        .bankno = Trim(txtBankNo.Text)
        .custid = Trim(txtId.Text)
        .custitem = Trim(txtItem.Text)
        .custman = Trim(txtMan.Text)
        .custmng = Trim(txtMng.Text)
        .custnm = Trim(txtNm.Text)
        .custtype = Trim(txtType.Text)
        .faxno = Trim(txtFax.Text)
        .hpno = Trim(txtHp.Text)
        .postno = Trim(txtPost.Text)
        .remark = Trim(txtRemark.Text)
        .telno = Trim(txtTel.Text)
        .delfg = chkDelete.Value
        
        If .cfSave Then
            lblNo.Caption = .custcd
            cmdDelete.Enabled = True
            
            Call psListRefresh
        End If
    End With
    MousePointer = vbDefault
    Exit Sub
    
ErrcmdSave:
    MsgBox Err.Description, vbCritical
    MousePointer = vbDefault
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    Call gsEnterEsc_KeyPress(Me, KeyAscii, Me.Count)

End Sub

Private Sub Form_Load()

    Set cCst = New clsMstCust
    
    Me.KeyPreview = True
    Me.Show
    
    Call psFieldLengthDefine
    Call cmdClear_Click
    Call psListRefresh

End Sub

Private Sub lblDelfg_Click()

End Sub

Private Sub spList_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim sData As Variant

    If Row > 0 And Col > 0 Then
        spList.GetText 1, Row, sData
        If Val(sData) > 0 Then
            gSql = "select * from mstCUST where custcd = " & sData
            With cDb.cfRecordSet(gSql)
                If .State = adStateOpen Then
                    If Not .EOF Then
                        lblNo.Caption = Val("" & .Fields("custcd").Value)
                        txtAddr1.Text = "" & .Fields("addr1").Value
                        txtAddr2.Text = "" & .Fields("addr2").Value
                        txtBank.Text = "" & .Fields("banknm").Value
                        txtBankNo.Text = "" & .Fields("bankno").Value
                        txtId.Text = "" & .Fields("custid").Value
                        txtItem.Text = "" & .Fields("custitem").Value
                        txtMan.Text = "" & .Fields("custman").Value
                        txtMng.Text = "" & .Fields("custmng").Value
                        txtNm.Text = "" & .Fields("custnm").Value
                        txtType.Text = "" & .Fields("custtype").Value
                        txtFax.Text = "" & .Fields("faxno").Value
                        txtHp.Text = "" & .Fields("hpno").Value
                        txtPost.Text = "" & .Fields("postno").Value
                        txtRemark.Text = "" & .Fields("remark").Value
                        txtTel.Text = "" & .Fields("telno").Value
                        lblWrtdt.Caption = "" & .Fields("wrtdt").Value
                        lblModdt.Caption = "" & .Fields("moddt").Value
                        chkDelete.Value = Val("" & .Fields("delfg").Value)
                        
                        cmdSave.Enabled = True
                        cmdDelete.Enabled = True
                    End If
                    .Close
                End If
            End With
        End If
    End If
    
End Sub

Private Sub txtNm_LostFocus()

    cmdSave.Enabled = Len(Trim(txtNm.Text)) > 0
    
End Sub
