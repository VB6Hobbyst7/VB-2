VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#8.0#0"; "FPSPRU80.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{1C203F10-95AD-11D0-A84B-00A0247B735B}#1.0#0"; "SSTree.ocx"
Begin VB.Form frm발주서일반 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "발주서작성(일반구매)"
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   18915
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
   ScaleHeight     =   9345
   ScaleWidth      =   18915
   ShowInTaskbar   =   0   'False
   Begin SSSplitter.SSSplitter splMain 
      Height          =   9345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18915
      _ExtentX        =   33364
      _ExtentY        =   16484
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   7
      SplitterBarJoinStyle=   0
      SplitterResizeStyle=   1
      SplitterBarAppearance=   1
      Locked          =   -1  'True
      PaneTree        =   "frm발주서일반.frx":0000
      Begin Threed.SSPanel SSPanel6 
         Height          =   360
         Left            =   3015
         TabIndex        =   13
         Top             =   8955
         Width           =   15870
         _ExtentX        =   27993
         _ExtentY        =   635
         _Version        =   262144
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "※ 입고상태의 자료는 삭제 하실수 없습니다 !!!"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin SSActiveTreeView.SSTree trvStkList 
         Height          =   8565
         Left            =   30
         TabIndex        =   10
         Top             =   750
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   15108
         _Version        =   65538
         Appearance      =   0
         Indentation     =   569.764
         PictureBackgroundUseMask=   0   'False
         HasFont         =   0   'False
         HasMouseIcon    =   0   'False
         HasPictureBackground=   0   'False
         ImageList       =   "<None>"
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1260
         Left            =   15315
         TabIndex        =   6
         Top             =   30
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   2223
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSCommand cmdSave 
            Height          =   420
            Left            =   120
            TabIndex        =   7
            Top             =   690
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
            Left            =   1230
            TabIndex        =   8
            Top             =   690
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
         Begin Threed.SSCommand cmdClose 
            Height          =   420
            Left            =   2340
            TabIndex        =   9
            Top             =   120
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
         Begin Threed.SSCommand cmdClear 
            Height          =   420
            Left            =   120
            TabIndex        =   11
            Top             =   120
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
         Begin Threed.SSCommand cmdPrint 
            Height          =   420
            Left            =   1230
            TabIndex        =   26
            Top             =   120
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
            Caption         =   "발주서출력"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand cmdOrdDelete 
            Height          =   420
            Left            =   2340
            TabIndex        =   36
            Top             =   690
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
            Caption         =   "발주서삭제"
            ButtonStyle     =   2
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1260
         Left            =   3015
         TabIndex        =   2
         Top             =   30
         Width           =   12195
         _ExtentX        =   21511
         _ExtentY        =   2223
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtRemark 
            Height          =   300
            Left            =   1320
            TabIndex        =   22
            Text            =   "Text1"
            Top             =   810
            Width           =   10695
         End
         Begin VB.TextBox txtNo 
            Alignment       =   2  '가운데 맞춤
            Height          =   300
            Left            =   2220
            TabIndex        =   18
            Text            =   "201206"
            Top             =   150
            Width           =   435
         End
         Begin VB.TextBox txtYm 
            Height          =   300
            Left            =   1320
            TabIndex        =   17
            Text            =   "201206"
            Top             =   150
            Width           =   705
         End
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   300
            Left            =   1320
            TabIndex        =   4
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            Format          =   21430273
            CurrentDate     =   41078
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   300
            Left            =   90
            TabIndex        =   3
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "발주일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   300
            Left            =   2700
            TabIndex        =   5
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "구매업체"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel lblSumAmt 
            Height          =   300
            Left            =   10350
            TabIndex        =   14
            Top             =   480
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   529
            _Version        =   262144
            ForeColor       =   255
            BackColor       =   -2147483624
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "발주번호 : 201207-1"
            BevelOuter      =   1
            Alignment       =   4
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   300
            Left            =   9120
            TabIndex        =   15
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "합계금액"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   300
            Left            =   90
            TabIndex        =   16
            Top             =   150
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "발주번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel9 
            Height          =   90
            Left            =   2040
            TabIndex        =   19
            Top             =   240
            Width           =   165
            _ExtentX        =   291
            _ExtentY        =   159
            _Version        =   262144
            BackColor       =   0
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel lblOrdType 
            Height          =   300
            Left            =   7830
            TabIndex        =   20
            Top             =   150
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   529
            _Version        =   262144
            BackColor       =   -2147483624
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "발주번호 : 201207-1"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel11 
            Height          =   300
            Left            =   6600
            TabIndex        =   21
            Top             =   150
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "발주구분"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel12 
            Height          =   300
            Left            =   90
            TabIndex        =   23
            Top             =   810
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "비고사항"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel lblUserNm 
            Height          =   300
            Left            =   10350
            TabIndex        =   24
            Top             =   150
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   529
            _Version        =   262144
            BackColor       =   -2147483624
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "발주번호 : 201207-1"
            BevelOuter      =   1
            Alignment       =   4
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel14 
            Height          =   300
            Left            =   9120
            TabIndex        =   25
            Top             =   150
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "사용자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin PVCOMBOLibCtl.PVComboBox cboCust 
            Height          =   300
            Left            =   3930
            TabIndex        =   27
            Top             =   480
            Width           =   975
            _Version        =   524288
            _cx             =   1720
            _cy             =   529
            Appearance      =   1
            Enabled         =   -1  'True
            BackColor       =   16777215
            ForeColor       =   0
            Locked          =   0   'False
            Style           =   2
            Sorted          =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowPictures    =   0   'False
            ColumnHeaders   =   -1  'True
            PrimaryColumn   =   0
            VisibleItems    =   10
            ColumnHeaderHeight=   20
            ListMember      =   ""
            ColumnHeaderForeColor=   0
            ColumnHeaderBackColor=   14215660
            SelectedForeColor=   16777215
            SelectedBackColor=   12937777
            AlternateBackColor=   16777215
            ItemLabelStyle  =   1
            ItemLabelType   =   0
            ItemLabelWidth  =   40
            ItemLabelForeColor=   0
            ItemLabelBackColor=   14215660
            ColumnHeaderStyle=   1
            VerticalGridLines=   0   'False
            HorizontalGridLines=   0   'False
            ColumnResize    =   0   'False
            ItemLabelResize =   0   'False
            AllowDBAutoConfig=   -1  'True
            GridLineColor   =   13421772
            List            =   ""
            NullString      =   "[NULL]"
            DropShadow      =   -1  'True
            Text            =   ""
            SortOnColumnHeaderClick=   0   'False
            DropEffect      =   1
            ColumnCount     =   2
            Column0.Heading =   "업체코드"
            Column0.Width   =   40
            Column0.Alignment=   0
            Column0.Hidden  =   0   'False
            Column0.Name    =   ""
            Column0.Format  =   ""
            Column0.Bound   =   0   'False
            Column0.Locked  =   0   'False
            Column0.HeaderAlignment=   0
            Column1.Heading =   "업체명"
            Column1.Width   =   80
            Column1.Alignment=   0
            Column1.Hidden  =   0   'False
            Column1.Name    =   ""
            Column1.Format  =   ""
            Column1.Bound   =   0   'False
            Column1.Locked  =   0   'False
            Column1.HeaderAlignment=   0
            SortKey1.Column =   -1
            SortKey1.Ascending=   -1  'True
            SortKey1.CaseInsensitive=   -1  'True
            SortKey2.Column =   -1
            SortKey2.Ascending=   -1  'True
            SortKey2.CaseInsensitive=   -1  'True
            SortKey3.Column =   -1
            SortKey3.Ascending=   -1  'True
            SortKey3.CaseInsensitive=   -1  'True
            BoundColumn     =   ""
            Border          =   -1  'True
            VertAlign       =   1
            Format          =   ""
         End
         Begin Threed.SSPanel lblCustNm 
            Height          =   300
            Left            =   4920
            TabIndex        =   28
            Top             =   480
            Width           =   4155
            _ExtentX        =   7329
            _ExtentY        =   529
            _Version        =   262144
            BackColor       =   -2147483624
            Caption         =   "업체번호"
            BevelOuter      =   1
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSCommand cmdOrdList 
            Height          =   300
            Left            =   2670
            TabIndex        =   29
            Top             =   150
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   529
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
            Caption         =   "..."
            ButtonStyle     =   2
         End
         Begin Threed.SSPanel SSPanel10 
            Height          =   300
            Left            =   3060
            TabIndex        =   30
            Top             =   150
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   " 기간 :"
            BevelOuter      =   1
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.OptionButton optMon6 
               Caption         =   "6개월"
               Height          =   255
               Left            =   2610
               TabIndex        =   33
               Top             =   30
               Width           =   855
            End
            Begin VB.OptionButton optMon3 
               Caption         =   "3개월"
               Height          =   255
               Left            =   1650
               TabIndex        =   32
               Top             =   30
               Width           =   855
            End
            Begin VB.OptionButton optMon1 
               Caption         =   "1개월"
               Height          =   255
               Left            =   690
               TabIndex        =   31
               Top             =   30
               Width           =   855
            End
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   615
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   2880
         _ExtentX        =   5080
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
         Caption         =   "▒ 물품현황 ▒"
         BevelOuter      =   1
         BevelInner      =   2
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin FPUSpreadADO.fpSpread spList 
         Height          =   7455
         Left            =   3015
         TabIndex        =   12
         Top             =   1395
         Width           =   15870
         _Version        =   524288
         _ExtentX        =   27993
         _ExtentY        =   13150
         _StockProps     =   64
         EditEnterAction =   5
         EditModeReplace =   -1  'True
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
         MaxCols         =   12
         SpreadDesigner  =   "frm발주서일반.frx":00D2
         UserResize      =   0
         AppearanceStyle =   0
      End
   End
   Begin Threed.SSPanel pnlOrdList 
      Height          =   5295
      Left            =   5700
      TabIndex        =   34
      Top             =   510
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   9340
      _Version        =   262144
      BackColor       =   8438015
      BevelWidth      =   2
      BorderWidth     =   2
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin FPUSpreadADO.fpSpread spOrdList 
         Height          =   5040
         Left            =   120
         TabIndex        =   35
         Top             =   120
         Width           =   6960
         _Version        =   524288
         _ExtentX        =   12277
         _ExtentY        =   8890
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
         GrayAreaBackColor=   -2147483634
         MaxCols         =   6
         OperationMode   =   3
         RowHeaderDisplay=   0
         SpreadDesigner  =   "frm발주서일반.frx":108A
         UserResize      =   0
      End
   End
End
Attribute VB_Name = "frm발주서일반"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub psListRefresh()
Dim sRow As Long
    
    gSql = "select a.*, b.custnm, c.usernm from ordH a with (nolock), mstCUST b, mstUSER c " & _
           "  where a.ordym = '" & Trim(txtYm.Text) & "' and a.ordno = " & Val(txtNo.Text) & " and a.custcd *= b.custcd and a.usercd = c.usercd"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                txtYm.Locked = True
                txtNo.Locked = True
                
                lblOrdType.Caption = gOrderType(Val("" & .Fields("ordtype").Value))
                lblOrdType.Tag = Val("" & .Fields("ordtype").Value)
                If Val("" & .Fields("stat").Value) = gOrderStsEnd Then
                    lblOrdType.Caption = lblOrdType.Caption & "/" & gOrderStat(Val("" & .Fields("stat").Value))
                    lblOrdType.ForeColor = vbRed
                Else
                    lblOrdType.ForeColor = vbBlack
                End If
                
                lblUserNm.Caption = "" & .Fields("usernm").Value
                lblSumAmt.Caption = Format(.Fields("ordamt").Value, "#,##0")
                dtpDate.Value = "" & .Fields("orddt").Value
                cboCust.Text = Val("" & .Fields("custcd").Value)
                lblCustNm.Caption = "" & .Fields("custnm").Value
                txtRemark.Text = "" & .Fields("remark").Value
                
                gSql = "select a.*, b.stknm, b.stkspec, b.buyunit from ordL a with (nolock), mstSTK b" & _
                       " where a.ordym = '" & .Fields("ordym").Value & "' and a.ordno = " & .Fields("ordno").Value & _
                       "   and a.stkcd = b.stkcd order by a.ordseq"
                With cDb.cfRecordSet(gSql)
                    If .State = adStateOpen Then
                        If Not .EOF Then
                            Call gsSpreadClear(spList, .RecordCount + 1000, True)
                            While (Not .EOF)
                                sRow = sRow + 1
                                
                                spList.SetText 1, sRow, ""
                                spList.SetText 2, sRow, "" & .Fields("stkcd").Value
                                spList.SetText 3, sRow, "" & .Fields("stknm").Value
                                spList.SetText 4, sRow, "" & .Fields("stkspec").Value
                                spList.SetText 5, sRow, "" & .Fields("buyunit").Value
                                spList.SetText 6, sRow, ""
                                spList.SetText 7, sRow, Val("" & .Fields("amt").Value)
                                spList.SetText 8, sRow, Val("" & .Fields("qty").Value)
                                spList.SetText 9, sRow, Val("" & .Fields("sumamt").Value)
                                spList.SetText 10, sRow, "" & .Fields("lastdt").Value
                                spList.SetText 11, sRow, "" & .Fields("remark").Value
                                spList.SetText 12, sRow, Val("" & .Fields("ordseq").Value)
                                
                                .MoveNext
                            Wend
                            
                            cmdDelete.Enabled = True
                            cmdSave.Enabled = True
                            cmdPrint.Enabled = True
                            cmdOrdDelete.Enabled = True
                        Else
                            Call gsSpreadClear(spList, , True)
                            
                            cmdDelete.Enabled = False
                            cmdSave.Enabled = False
                            cmdPrint.Enabled = False
                            cmdOrdDelete.Enabled = False
                        End If
                        .Close
                    End If
                End With
            Else
                txtYm.Locked = False
                txtNo.Locked = False
            End If
            .Close
        End If
    End With
    
End Sub

Private Function pfDataProcess(ByVal brJob As Boolean) As Boolean
Dim cOrdL As clsOrdList
Dim sRow As Long, sData As Variant, sReturn As Boolean, sCode As Long, sSeq As Integer
    
    sReturn = True
    
    Set cOrdL = New clsOrdList
    With spList
        For sRow = 1 To .MaxRows
            .GetText .MaxCols, sRow, sData:     sSeq = Val(sData)
            .GetText 2, sRow, sData:            sCode = Val(sData)
            .GetText 1, sRow, sData
            If Val(sData) > 0 And sCode > 0 Then
                If brJob Then
                    cOrdL.ordym = Trim(txtYm.Text)
                    cOrdL.ordno = Val(txtNo.Text)
                    cOrdL.stkcd = sCode
                    cOrdL.ordseq = sSeq
                    .GetText 7, sRow, sData:            cOrdL.amt = Val(sData)
                    .GetText 8, sRow, sData:            cOrdL.qty = Trim(sData)
                    .GetText 9, sRow, sData:            cOrdL.sumamt = Trim(sData)
                    .GetText 10, sRow, sData:           cOrdL.lastdt = Trim(sData)
                    .GetText 11, sRow, sData:           cOrdL.remark = Trim(sData)
                    sReturn = cOrdL.cfSave
                    If sReturn Then
                        .SetText .MaxCols, sRow, cOrdL.ordseq
                    End If
                Else
                    If sSeq > 0 Then
                        Call cDb.csBegin
                        sReturn = cOrdL.cfDelete(Trim(txtYm.Text), Val(txtNo.Text), sSeq)
                        If sReturn Then
                            Call cDb.csCommit
                        Else
                            Call cDb.csRollback
                        End If
                    End If
                End If
                
                If sReturn = False Then
                    Exit For
                Else
                    .SetText 1, sRow, ""
                End If
            End If
        Next sRow
    End With
    
    pfDataProcess = sReturn
    
    If sReturn Then
        Call psListRefresh
    End If

End Function

Private Sub cboCust_Click()

    If cboCust.ListIndex > 0 Then
        lblCustNm.Caption = cboCust.SubItem(cboCust.ListIndex, 1)
        
        cmdSave.Enabled = True
    Else
        lblCustNm.Caption = ""
        cmdSave.Enabled = False
    End If

End Sub

Private Sub cmdClear_Click()
    
    pnlOrdList.Visible = False
    
    Call gsSpreadClear(spList, , True)
    Call gsSetStkTree(trvStkList)

    dtpDate.Value = gfSystemDate
    Call gsSetCustComboPV(cboCust)
    
    txtYm.Text = ""
    txtNo.Text = ""
    lblOrdType.Caption = ""
    lblUserNm.Caption = ""
    lblSumAmt.Caption = ""
    lblCustNm.Caption = ""
    txtRemark.Text = ""
    
    txtYm.Text = Format(dtpDate.Value, "yyyyMM")
    
    cmdSave.Enabled = False
    cmdDelete.Enabled = False
    cmdPrint.Enabled = False
    cmdOrdDelete.Enabled = False

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdDelete_Click()
Dim cOrd As clsOrdHead, sReturn As Boolean, sAmt As Currency

    MousePointer = vbHourglass
    If MsgBox("선택하신 요청자료를 삭제하시겠습니까 ?", vbYesNo + vbQuestion) = vbYes Then
        Set cOrd = New clsOrdHead
        sReturn = pfDataProcess(False)
        
        If sReturn Then
            sReturn = cOrd.cfAmtSumUpdate(txtYm.Text, Val(txtNo.Text), sAmt)
        End If
    
        If sReturn Then
            lblSumAmt.Caption = Format(sAmt, "#,##0")
            Call cDb.csCommit
        Else
            Call cDb.csRollback
        End If
    End If
    MousePointer = vbDefault

End Sub

Private Sub cmdOrdDelete_Click()
Dim cOrd As clsOrdHead, sReturn As Boolean

    MousePointer = vbHourglass
    If MsgBox("선택하신 발주서 전체를 삭제하시겠습니까 ?", vbYesNo + vbQuestion) = vbYes Then
        Set cOrd = New clsOrdHead
        
        gSql = "select ordym from buyL with (nolock) where ordym = '" & txtYm.Text & "' and ordno = " & Val(txtNo.Text)
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                If .EOF Then
                    sReturn = True
                Else
                    sReturn = False
                    
                    MsgBox "발주서 품목중 입고가 진행된 자료가 있습니다.!", vbCritical
                End If
                .Close
            End If
        End With
        
        If sReturn Then
            Call cDb.csBegin
            gSql = "update reqL set ordym = '', ordno = 0, ordseq = 0, stat = " & gReqStatWrt & _
                   " where ordym = '" & txtYm.Text & "' and ordno = " & Val(txtNo.Text)
            sReturn = cDb.cfExecute(gSql)
            
            If sReturn Then
                sReturn = cOrd.cfDelete(txtYm.Text, Val(txtNo.Text))
            End If
            
            If sReturn Then
                Call cDb.csCommit
                
                Call cmdClear_Click
            Else
                Call cDb.csRollback
            End If
        End If
    End If
    MousePointer = vbDefault
    
End Sub

Private Sub cmdOrdList_Click()
Dim sRow As Long, sPrevDate As String, sMon As Integer, sDate As String
    
    sDate = Format(gfSystemDate, "yyyy-MM-dd")
    
    If optMon1.Value Then
        sMon = -1
    ElseIf optMon3.Value Then
        sMon = -3
    ElseIf optMon6.Value Then
        sMon = -6
    Else
        sMon = -1
    End If
    
    sPrevDate = Format(DateAdd("m", sMon, sDate), "yyyy-MM-dd")
    
    gSql = "select a.*, b.custnm, c.usernm from ordH a with (nolock), mstCUST b, mstUSER c " & _
           " where a.orddt between '" & sPrevDate & "' and '" & sDate & "'" & _
           "   and a.custcd *= b.custcd and a.usercd = c.usercd order by a.ordym DESC, a.ordno DESC"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                Call gsSpreadClear(spOrdList, .RecordCount)
                While (Not .EOF)
                    sRow = sRow + 1
                    
                    spOrdList.SetText 1, sRow, "" & .Fields("ordym").Value & "-" & .Fields("ordno").Value
                    spOrdList.SetText 2, sRow, "" & .Fields("orddt").Value
                    spOrdList.SetText 3, sRow, "" & .Fields("custnm").Value
                    spOrdList.SetText 4, sRow, "" & .Fields("usernm").Value
                    
                    spOrdList.SetText 5, sRow, "" & .Fields("ordym").Value
                    spOrdList.SetText 6, sRow, Val("" & .Fields("ordno").Value)
                    
                    .MoveNext
                Wend
                
                pnlOrdList.Visible = True
                pnlOrdList.ZOrder 0
                splMain.Enabled = False
                
                spOrdList.SetFocus
            Else
                MsgBox "최근 등록된 발주서가 없습니다.!", vbCritical
            End If
            .Close
        End If
    End With
    
End Sub

Private Sub cmdSave_Click()
Dim cOrd As clsOrdHead, sReturn As Boolean, sAmt As Currency

    MousePointer = vbHourglass
    Set cOrd = New clsOrdHead
    
    Call cDb.csBegin
    With cOrd
        .ordym = Trim(txtYm.Text)
        .ordno = Val(txtNo.Text)
        .custcd = Val(cboCust.Text)
'        .ordamt = Val(Str(lblSumAmt.Caption))
        .orddt = Format(dtpDate.Value, "yyyy-MM-dd")
        .stat = gOrderStsWrt
        .ordtype = Val(lblOrdType.Tag)
        .remark = Trim(txtRemark.Text)
        
        sReturn = .cfSave
        If sReturn Then
            sReturn = pfDataProcess(True)
        End If
        
        If sReturn Then
            sReturn = .cfAmtSumUpdate(txtYm.Text, Val(txtNo.Text), sAmt)
        End If
    End With
    
    If sReturn Then
        txtYm.Locked = True
        txtNo.Locked = True
        
        txtNo.Text = cOrd.ordno
        
        lblSumAmt.Caption = Format(sAmt, "#,##0")
        Call cDb.csCommit
    Else
        Call cDb.csRollback
    End If
    MousePointer = vbDefault
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape And pnlOrdList.Visible Then
        pnlOrdList.Visible = False
        splMain.Enabled = True
    End If
    
End Sub

Private Sub Form_Load()
Dim sCol As Integer

    Me.Width = 19000
    Me.Height = 10120
    
    Me.KeyPreview = True
    Me.Show
    
    ' 스프레스 Sort 설정
    With spOrdList
        .UserColAction = UserColActionSort
        For sCol = 1 To .MaxCols
            .ColUserSortIndicator(sCol) = ColUserSortIndicatorAscending
        Next sCol
    End With
    
    Call cmdClear_Click
    
    optMon1.Value = True

End Sub

Private Sub spList_Change(ByVal Col As Long, ByVal Row As Long)
Dim sData As Variant, sDate As String, sAmt As Currency, sQty As Single

    If Row > 0 And Col > 1 Then
        spList.SetText 1, Row, "1"
        
        Select Case Col
            Case 2
                    sDate = Format(dtpDate.Value, "yyyy-MM-dd")
                    spList.GetText 2, Row, sData
                    gSql = "select stkcd, stknm, stkspec, buyunit, buyday, buyamt from mstSTK where stkcd = " & Val(sData)
                    With cDb.cfRecordSet(gSql)
                        If .State = adStateOpen Then
                            If Not .EOF Then
                                spList.SetText 3, Row, "" & .Fields("stknm").Value
                                spList.SetText 4, Row, "" & .Fields("stkspec").Value
                                spList.SetText 5, Row, "" & .Fields("buyunit").Value
                                spList.SetText 6, Row, gfPresentStkRmd(Val(sData), sDate, False)
                                spList.SetText 7, Row, Val("" & .Fields("buyamt").Value)
                                spList.SetText 10, Row, DateAdd("d", Val("" & .Fields("buyday").Value), sDate)
                                
                                spList.Row = Row
                                spList.Col = 7
                                spList.Action = ActionActiveCell
                            Else
                                MsgBox "등록되지 않은 물품입니다.!", vbCritical
                                
                                spList.SetText 2, Row, ""
                                spList.SetText 3, Row, ""
                                spList.SetText 4, Row, ""
                                spList.SetText 5, Row, ""
                                spList.SetText 6, Row, ""
                                spList.SetText 7, Row, ""
                                spList.SetText 10, Row, ""
                            
                                spList.Row = Row
                                spList.Col = 2
                                spList.Action = ActionActiveCell
                                spList.SetFocus
                            End If
                            .Close
                        End If
                    End With
            Case 7, 8
                    With spList
                        .GetText 7, Row, sData:     sAmt = Val(sData)
                        .GetText 8, Row, sData:     sQty = Val(sData)
                        .SetText 9, Row, sAmt * sQty
                    End With
        End Select
    End If

End Sub

Private Sub spOrdList_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim sData As Variant
 
    If Row > 0 And Col > 0 Then
        With spOrdList
            .GetText 5, Row, sData:     txtYm.Text = Trim(sData)
            .GetText 6, Row, sData:     txtNo.Text = Val(sData)
            
            Call psListRefresh
            
            pnlOrdList.Visible = False
            splMain.Enabled = True
        End With
    End If
    
End Sub

Private Sub trvStkList_Collapse(Node As SSActiveTreeView.SSNode)

    If Node.Level = 1 Then Node.Image = "close"

End Sub

Private Sub trvStkList_DblClick()

    If trvStkList.SelectedNodes.Item(1).Level > 1 Then
        Call psStkAppendCheck
    End If

End Sub

Private Sub trvStkList_Expand(Node As SSActiveTreeView.SSNode)

    If Node.Level = 1 Then Node.Image = "open"

End Sub

Private Sub psStkAppendCheck()
Dim sRow As Long, sData As Variant

    With spList
        For sRow = 1 To .MaxRows
            .GetText 2, sRow, sData
            If Len(sData) = 0 Then
                .SetText 2, sRow, trvStkList.SelectedItem.Key
                Call spList_Change(2, sRow)
                
                spList.Row = sRow
                spList.Col = 8
                spList.Action = ActionActiveCell
                Exit For
'            ElseIf Trim(sData) = trvStkList.SelectedItem.Key Then
'                MsgBox "등록된 품번입니다.!", vbCritical
'                Exit For
            End If
        Next sRow
    End With

End Sub

Private Sub txtNo_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Call psListRefresh
    End If

End Sub
