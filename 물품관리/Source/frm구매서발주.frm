VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#8.0#0"; "FPSPRU80.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm구매서발주 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "구매서작성(일반구매)"
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
      TabIndex        =   3
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
      PaneTree        =   "frm구매서발주.frx":0000
      Begin Threed.SSPanel SSPanel5 
         Height          =   885
         Left            =   15345
         TabIndex        =   7
         Top             =   30
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   1561
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSCommand cmdSave 
            Height          =   420
            Left            =   1230
            TabIndex        =   8
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
            Caption         =   "저장(&S)"
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
            TabIndex        =   10
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
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   885
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   1561
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   300
            Left            =   7860
            TabIndex        =   2
            Top             =   450
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            Format          =   21430273
            CurrentDate     =   41078
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   300
            Left            =   6630
            TabIndex        =   5
            Top             =   450
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "구매일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   300
            Left            =   90
            TabIndex        =   6
            Top             =   120
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
            Left            =   13410
            TabIndex        =   11
            Top             =   120
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
            Left            =   12180
            TabIndex        =   12
            Top             =   120
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "합계금액"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin PVCOMBOLibCtl.PVComboBox cboCust 
            Height          =   300
            Left            =   1320
            TabIndex        =   0
            Top             =   120
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
            Left            =   2310
            TabIndex        =   13
            Top             =   120
            Width           =   4275
            _ExtentX        =   7541
            _ExtentY        =   529
            _Version        =   262144
            BackColor       =   -2147483624
            Caption         =   "업체번호"
            BevelOuter      =   1
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel lblOrdNo 
            Height          =   300
            Left            =   1320
            TabIndex        =   14
            Top             =   450
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   529
            _Version        =   262144
            ForeColor       =   255
            BackColor       =   -2147483634
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
         Begin Threed.SSPanel lblOrdYm 
            Height          =   300
            Left            =   1320
            TabIndex        =   15
            Top             =   750
            Visible         =   0   'False
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   529
            _Version        =   262144
            ForeColor       =   255
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
         Begin Threed.SSPanel lblOrdSeq 
            Height          =   300
            Left            =   2250
            TabIndex        =   16
            Top             =   750
            Visible         =   0   'False
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   529
            _Version        =   262144
            ForeColor       =   255
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
         Begin Threed.SSPanel SSPanel6 
            Height          =   300
            Left            =   90
            TabIndex        =   17
            Top             =   450
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "발주번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSCommand cmdOrdList 
            Height          =   300
            Left            =   2700
            TabIndex        =   1
            Top             =   450
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
            Left            =   3090
            TabIndex        =   18
            Top             =   450
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   " 기간 :"
            BevelOuter      =   1
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.OptionButton optMon1 
               Caption         =   "1개월"
               Height          =   255
               Left            =   690
               TabIndex        =   21
               Top             =   30
               Width           =   855
            End
            Begin VB.OptionButton optMon3 
               Caption         =   "3개월"
               Height          =   255
               Left            =   1650
               TabIndex        =   20
               Top             =   30
               Width           =   855
            End
            Begin VB.OptionButton optMon6 
               Caption         =   "6개월"
               Height          =   255
               Left            =   2610
               TabIndex        =   19
               Top             =   30
               Width           =   855
            End
         End
         Begin Threed.SSPanel lblOrdType 
            Height          =   300
            Left            =   7860
            TabIndex        =   24
            Top             =   120
            Width           =   1305
            _ExtentX        =   2302
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
            Left            =   6630
            TabIndex        =   25
            Top             =   120
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "발주구분"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel lblUserNm 
            Height          =   300
            Left            =   10470
            TabIndex        =   26
            Top             =   120
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
            Left            =   9240
            TabIndex        =   27
            Top             =   120
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "발주등록자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel lblRemark 
            Height          =   300
            Left            =   10470
            TabIndex        =   28
            Top             =   450
            Width           =   4605
            _ExtentX        =   8123
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
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   300
            Left            =   9240
            TabIndex        =   29
            Top             =   450
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "발주비고사항"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
      Begin FPUSpreadADO.fpSpread spList 
         Height          =   7815
         Left            =   30
         TabIndex        =   30
         Top             =   1020
         Width           =   18855
         _Version        =   524288
         _ExtentX        =   33258
         _ExtentY        =   13785
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
         MaxCols         =   15
         SpreadDesigner  =   "frm구매서발주.frx":0092
         UserResize      =   0
         AppearanceStyle =   0
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   375
         Left            =   30
         TabIndex        =   31
         Top             =   8940
         Width           =   18855
         _ExtentX        =   33258
         _ExtentY        =   661
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
         Caption         =   "※ 발주서 기준에 의한 입고처리만 가능합니다. 수정 및 삭제는 [발주서 입고]메뉴에서 작업하시기 바랍니다. !!!"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin Threed.SSPanel pnlOrdList 
      Height          =   5295
      Left            =   2340
      TabIndex        =   22
      Top             =   780
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
         TabIndex        =   23
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
         SpreadDesigner  =   "frm구매서발주.frx":116B
         UserResize      =   0
      End
   End
End
Attribute VB_Name = "frm구매서발주"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub psListRefresh()
Dim sRow As Long, sDate As String, sAmt As Currency

    sDate = Format(dtpDate.Value, "yyyy-MM-dd")
    gSql = "select a.ordtype, a.ordamt, a.remark, b.usernm from ordH a with (nolock), mstUSER b " & _
           " where a.ordym = '" & lblOrdYm.Caption & "' and ordno = " & Val(lblOrdSeq.Caption) & _
           "   and a.usercd = b.usercd"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                lblOrdType.Caption = gOrderType(Val("" & .Fields("ordtype").Value))
                lblUserNm.Caption = "" & .Fields("usernm").Value
                lblSumAmt.Caption = Format(Val("" & .Fields("ordamt").Value), "#,##0")
                lblRemark.Caption = "" & .Fields("remark").Value
            End If
            .Close
        End If
    End With
    
    gSql = "select a.ordseq, a.stkcd, a.qty, a.amt as ordamt, a.inqty, b.stknm, b.stkspec, b.buyunit, b.buyioqty from ordL a with (nolock), mstSTK b " & _
           " where a.ordym = '" & lblOrdYm.Caption & "' and a.ordno = " & Val(lblOrdSeq.Caption) & " and a.stkcd = b.stkcd " & _
           " order by a.ordseq"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                Call gsSpreadClear(spList, .RecordCount, True)
                While (Not .EOF)
                    sRow = sRow + 1
                    
                    spList.SetText 1, sRow, ""
                    spList.SetText 2, sRow, "" & .Fields("stkcd").Value
                    spList.SetText 3, sRow, "" & .Fields("stknm").Value
                    spList.SetText 4, sRow, "" & .Fields("stkspec").Value
                    spList.SetText 5, sRow, "" & .Fields("buyunit").Value
                    spList.SetText 6, sRow, Val("" & .Fields("qty").Value)
                    spList.SetText 7, sRow, Val("" & .Fields("qty").Value) - Val("" & .Fields("inqty").Value)
                    spList.SetText 8, sRow, Val("" & .Fields("buyioqty").Value)
                    spList.SetText 9, sRow, Val("" & .Fields("ordamt").Value)
                    spList.SetText 10, sRow, ""
                    spList.SetText 11, sRow, ""
                    spList.SetText 12, sRow, ""
                    spList.SetText 13, sRow, ""
                    spList.SetText 14, sRow, Val("" & .Fields("ordseq").Value)
                    spList.SetText 15, sRow, ""
                    
                    .MoveNext
                Wend
                cmdSave.Enabled = True
            Else
                Call gsSpreadClear(spList, 0, True)
                cmdSave.Enabled = False
            End If
            .Close
        End If
    End With

End Sub

Private Sub cboCust_Click()

    If cboCust.ListIndex > 0 Then
        lblCustNm.Caption = cboCust.SubItem(cboCust.ListIndex, 1)
    Else
        lblCustNm.Caption = ""
    End If

End Sub

Private Sub cmdClear_Click()
    
    Call gsSpreadClear(spList, , True)
    pnlOrdList.Visible = False

    dtpDate.Value = gfSystemDate
    Call gsSetCustComboPV(cboCust)
    
    cboCust.Text = ""
    lblSumAmt.Caption = ""
    lblCustNm.Caption = ""
    lblOrdYm.Caption = ""
    lblOrdSeq.Caption = ""
    lblOrdNo.Caption = ""
    lblUserNm.Caption = ""
    lblOrdType.Caption = ""
    lblRemark.Caption = ""
    
    cmdSave.Enabled = False

End Sub

Private Sub cmdClose_Click()

    Unload Me

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
           " where a.orddt between '" & sPrevDate & "' and '" & sDate & "' and a.custcd = " & Val(cboCust.Text) & " and (stat < " & gOrderStsEnd & " or stat is null)" & _
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
Dim cBuy As clsBuyList
Dim sRow As Long, sData As Variant, sDate As String, sReturn As Boolean, sCode As Long, sSeq As Integer
    
    MousePointer = vbHourglass
    sReturn = True
    sDate = Format(dtpDate.Value, "yyyy-MM-dd")
    
    Set cBuy = New clsBuyList
    With spList
        For sRow = 1 To .MaxRows
            .GetText .MaxCols, sRow, sData:     sSeq = Val(sData)
            .GetText 2, sRow, sData:            sCode = Val(sData)
            .GetText 1, sRow, sData
            If Val(sData) > 0 And sCode > 0 Then
                Call cDb.csBegin
                cBuy.buydt = sDate
                cBuy.buyseq = sSeq
                cBuy.stkcd = sCode
                cBuy.custcd = Val(cboCust.Text)
                cBuy.usercd = gUserId
                cBuy.buytype = gBuyIoOrder
                cBuy.ordym = lblOrdYm.Caption
                cBuy.ordno = Val(lblOrdSeq.Caption)
                .GetText 8, sRow, sData:            cBuy.qtyrate = Val(sData)
                .GetText 9, sRow, sData:            cBuy.amt = Val(sData)
                .GetText 10, sRow, sData:           cBuy.buyqty = Trim(sData)
                .GetText 11, sRow, sData:           cBuy.sumamt = Trim(sData)
                .GetText 12, sRow, sData:           cBuy.maxdt = Trim(sData)
                .GetText 13, sRow, sData:           cBuy.makeno = Trim(sData)
                .GetText 14, sRow, sData:           cBuy.ordseq = Val(sData)
                sReturn = cBuy.cfSave
                If sReturn Then
                    .SetText .MaxCols, sRow, cBuy.buyseq
                End If
                
                If sReturn = False Then
                    Call cDb.csRollback
                    Exit For
                Else
                    Call cDb.csCommit
                    .SetText 1, sRow, ""
                End If
            End If
        Next sRow
    End With
    
    If sReturn Then
        Call psListRefresh
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
    
    Call cmdClear_Click

End Sub

Private Sub spList_Change(ByVal Col As Long, ByVal Row As Long)
Dim sData As Variant, sDate As String, sAmt As Currency, sQty As Single

    If Row > 0 And Col > 1 Then
        With spList
            .SetText 1, Row, "1"
        
            If Col = 9 Or Col = 10 Then
                .GetText 9, Row, sData:     sAmt = Val(sData)
                .GetText 10, Row, sData:    sQty = Val(sData)
                .SetText 11, Row, sAmt * sQty
            End If
        End With
    End If

End Sub

Private Sub spList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 2 Then
        spList.Row = spList.ActiveRow
        spList.OperationMode = OperationModeSingle
    End If
    
End Sub

Private Sub spList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim sPt As POINTAPI, sRet As Long, sCol As Integer, sRow As Integer, sData As Variant
    
    If Button = 2 Then
        hMenu = CreatePopupMenu()
        AppendMenu hMenu, MF_STRING, 1, "품목이중입고"
'        AppendMenu hMenu, MF_GRAYED Or MF_DISABLED, 2, "추가하기"
'        AppendMenu hMenu, MF_SEPARATOR, 3, ByVal 0&
'        AppendMenu hMenu, MF_CHECKED, 4, "이 프로그램은..."
        GetCursorPos sPt
        sRet = TrackPopupMenuEx(hMenu, TPM_LEFTALIGN Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, sPt.x, sPt.y, Me.HWnd, ByVal 0&)
        DestroyMenu hMenu
        
        If sRet = 1 Then
            spList.MaxRows = spList.MaxRows + 1
            spList.Row = spList.ActiveRow + 1
            spList.Action = ActionInsertRow
            Call gsSpreadClear(spList, spList.MaxRows, True, , True)
            
            sRow = spList.ActiveRow
            For sCol = 1 To spList.MaxCols - 1
                spList.GetText sCol, sRow, sData
                spList.SetText sCol, sRow + 1, sData
            Next sCol
        End If
        
        spList.OperationMode = OperationModeNormal
    End If

End Sub

Private Sub spOrdList_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim sData As Variant
 
    If Row > 0 And Col > 0 Then
        With spOrdList
            .GetText 5, Row, sData:     lblOrdYm.Caption = Trim(sData)
            .GetText 6, Row, sData:     lblOrdSeq.Caption = Val(sData)
            
            lblOrdNo.Caption = lblOrdYm.Caption & "-" & lblOrdSeq.Caption
            
            Call psListRefresh
            
            pnlOrdList.Visible = False
            splMain.Enabled = True
        End With
    End If
    
End Sub
