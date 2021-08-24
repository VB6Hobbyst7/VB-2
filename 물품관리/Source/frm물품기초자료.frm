VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#8.0#0"; "FPSPRU80.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{C2000000-FFFF-1100-8100-000000000001}#8.0#0"; "PVCurr.ocx"
Object = "{C2000000-FFFF-1100-8200-000000000001}#8.0#0"; "PVNum.ocx"
Begin VB.Form frm물품기초자료 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "물품기초자료"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14550
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
   ScaleHeight     =   7665
   ScaleWidth      =   14550
   ShowInTaskbar   =   0   'False
   Begin SSSplitter.SSSplitter splMain 
      Height          =   7665
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   14550
      _ExtentX        =   25665
      _ExtentY        =   13520
      _Version        =   262144
      AutoSize        =   1
      SplitterBarWidth=   7
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   1
      Locked          =   -1  'True
      PaneTree        =   "frm물품기초자료.frx":0000
      Begin FPUSpreadADO.fpSpread spList 
         Height          =   6210
         Left            =   6240
         TabIndex        =   18
         Top             =   750
         Width           =   8280
         _Version        =   524288
         _ExtentX        =   14605
         _ExtentY        =   10954
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
         MaxCols         =   4
         OperationMode   =   3
         SpreadDesigner  =   "frm물품기초자료.frx":00D2
         UserResize      =   0
         AppearanceStyle =   0
      End
      Begin Threed.SSPanel SSPanel26 
         Height          =   570
         Left            =   6240
         TabIndex        =   19
         Top             =   7065
         Width           =   8280
         _ExtentX        =   14605
         _ExtentY        =   1005
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   1380
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   150
            Width           =   5385
         End
         Begin Threed.SSPanel SSPanel27 
            Height          =   300
            Left            =   150
            TabIndex        =   21
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
            Caption         =   "검색"
            ButtonStyle     =   2
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   6210
         Left            =   30
         TabIndex        =   23
         Top             =   750
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   10954
         _Version        =   262144
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin PVNumericLib.PVNumeric numSafeQty 
            Height          =   300
            Left            =   4290
            TabIndex        =   12
            Top             =   3780
            Width           =   735
            _Version        =   524288
            _ExtentX        =   1296
            _ExtentY        =   529
            _StockProps     =   253
            Text            =   "0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            Alignment       =   1
            EditMode        =   0
            SpinButtons     =   0
            LimitValue      =   -1  'True
         End
         Begin PVNumericLib.PVNumeric numMinQty 
            Height          =   300
            Left            =   1410
            TabIndex        =   13
            Top             =   4110
            Width           =   1575
            _Version        =   524288
            _ExtentX        =   2778
            _ExtentY        =   529
            _StockProps     =   253
            Text            =   "0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            Alignment       =   1
            EditMode        =   0
            SpinButtons     =   0
            LimitValue      =   -1  'True
         End
         Begin PVCurrencyLib.PVCurrency numStdAmt 
            Height          =   300
            Left            =   1410
            TabIndex        =   14
            Top             =   4530
            Width           =   1575
            _Version        =   524288
            _ExtentX        =   2778
            _ExtentY        =   529
            _StockProps     =   253
            Text            =   "\0.0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            Alignment       =   2
            EditMode        =   0
            Symbol          =   "\"
            DecimalPlaces   =   "1"
            Value           =   0
         End
         Begin PVCurrencyLib.PVCurrency numBuyAmt 
            Height          =   300
            Left            =   4260
            TabIndex        =   15
            Top             =   4560
            Width           =   1575
            _Version        =   524288
            _ExtentX        =   2778
            _ExtentY        =   529
            _StockProps     =   253
            Text            =   "\0.0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            Alignment       =   2
            EditMode        =   0
            Symbol          =   "\"
            DecimalPlaces   =   "1"
            Value           =   0
         End
         Begin PVNumericLib.PVNumeric numRateQty 
            Height          =   300
            Left            =   1410
            TabIndex        =   7
            Top             =   2610
            Width           =   1575
            _Version        =   524288
            _ExtentX        =   2778
            _ExtentY        =   529
            _StockProps     =   253
            Text            =   "0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            EditMode        =   0
            SpinButtons     =   0
            LimitValue      =   -1  'True
         End
         Begin PVNumericLib.PVNumeric numBuyDay 
            Height          =   300
            Left            =   4290
            TabIndex        =   10
            Top             =   3360
            Width           =   735
            _Version        =   524288
            _ExtentX        =   1296
            _ExtentY        =   529
            _StockProps     =   253
            Text            =   "0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            Alignment       =   1
            EditMode        =   0
            SpinButtons     =   0
            LimitValue      =   -1  'True
         End
         Begin VB.CheckBox chkDelete 
            Caption         =   "삭제"
            Height          =   300
            Left            =   1560
            TabIndex        =   16
            Top             =   4980
            Width           =   975
         End
         Begin PVCOMBOLibCtl.PVComboBox cboKind 
            Height          =   300
            Left            =   1410
            TabIndex        =   2
            Top             =   1170
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
            Column0.Heading =   "분류코드"
            Column0.Width   =   30
            Column0.Alignment=   0
            Column0.Hidden  =   0   'False
            Column0.Name    =   ""
            Column0.Format  =   ""
            Column0.Bound   =   0   'False
            Column0.Locked  =   0   'False
            Column0.HeaderAlignment=   0
            Column1.Heading =   "분류명"
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
         Begin VB.ComboBox cboRmd 
            Height          =   300
            Left            =   1410
            Style           =   2  '드롭다운 목록
            TabIndex        =   11
            Top             =   3780
            Width           =   1575
         End
         Begin VB.ComboBox cboBuyType 
            Height          =   300
            Left            =   1410
            Style           =   2  '드롭다운 목록
            TabIndex        =   9
            Top             =   3360
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
         Begin VB.TextBox txtSpec 
            Height          =   300
            Left            =   1410
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   840
            Width           =   4455
         End
         Begin VB.TextBox txtIoUnit 
            Height          =   300
            Left            =   4290
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   2280
            Width           =   1575
         End
         Begin VB.TextBox txtMaker 
            Height          =   300
            Left            =   1410
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   1500
            Width           =   4455
         End
         Begin VB.TextBox txtBuyUnit 
            Height          =   300
            Left            =   1410
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   2280
            Width           =   1575
         End
         Begin VB.TextBox txtBarcode 
            Height          =   300
            Left            =   1410
            TabIndex        =   4
            Text            =   "123456789012345"
            Top             =   1830
            Width           =   4455
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   300
            Left            =   180
            TabIndex        =   24
            Top             =   180
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "품목코드"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel6 
            Height          =   300
            Left            =   180
            TabIndex        =   25
            Top             =   510
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "품 명"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   300
            Left            =   180
            TabIndex        =   26
            Top             =   840
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "규 격"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   300
            Left            =   3060
            TabIndex        =   27
            Top             =   2280
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "수불단위"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel9 
            Height          =   300
            Left            =   180
            TabIndex        =   28
            Top             =   1500
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "제조회사"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel10 
            Height          =   300
            Left            =   180
            TabIndex        =   29
            Top             =   2280
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "구매단위"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel12 
            Height          =   300
            Left            =   180
            TabIndex        =   30
            Top             =   1830
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "바코드번호"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel13 
            Height          =   300
            Left            =   180
            TabIndex        =   31
            Top             =   2610
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "구매:수불비"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel14 
            Height          =   300
            Left            =   180
            TabIndex        =   32
            Top             =   1170
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "물품분류"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel15 
            Height          =   300
            Left            =   180
            TabIndex        =   33
            Top             =   3030
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "주매입처"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel16 
            Height          =   300
            Left            =   3060
            TabIndex        =   34
            Top             =   3360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "구매기간"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel17 
            Height          =   300
            Left            =   180
            TabIndex        =   35
            Top             =   3780
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "재고관리"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel18 
            Height          =   300
            Left            =   3060
            TabIndex        =   36
            Top             =   3780
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "적정재고"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel19 
            Height          =   300
            Left            =   180
            TabIndex        =   37
            Top             =   4110
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "최소구매량"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel lblNo 
            Height          =   300
            Left            =   1410
            TabIndex        =   38
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
            Left            =   180
            TabIndex        =   39
            Top             =   3360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "구매유형"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel lblWrtdt 
            Height          =   300
            Left            =   1410
            TabIndex        =   40
            Top             =   5430
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
            TabIndex        =   41
            Top             =   5760
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   300
            Left            =   180
            TabIndex        =   49
            Top             =   4530
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "표준단가"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel20 
            Height          =   300
            Left            =   3030
            TabIndex        =   50
            Top             =   4560
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "구매단가"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel21 
            Height          =   300
            Left            =   180
            TabIndex        =   51
            Top             =   5430
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "등록일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel23 
            Height          =   300
            Left            =   180
            TabIndex        =   52
            Top             =   5760
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "수정일자"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel24 
            Height          =   300
            Left            =   3060
            TabIndex        =   53
            Top             =   2610
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "1(구매단위) : ?(수불단위)"
            BevelOuter      =   0
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel lblKindNm 
            Height          =   300
            Left            =   2400
            TabIndex        =   56
            Top             =   1170
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   529
            _Version        =   262144
            BackColor       =   -2147483624
            Caption         =   "업체번호"
            BevelOuter      =   1
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin PVCOMBOLibCtl.PVComboBox cboCust 
            Height          =   300
            Left            =   1410
            TabIndex        =   8
            Top             =   3030
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
            Left            =   2400
            TabIndex        =   57
            Top             =   3030
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   529
            _Version        =   262144
            BackColor       =   -2147483624
            Caption         =   "업체번호"
            BevelOuter      =   1
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel28 
            Height          =   300
            Left            =   180
            TabIndex        =   58
            Top             =   4980
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "사용안함"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel29 
            Height          =   300
            Left            =   5040
            TabIndex        =   59
            Top             =   3360
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "(일)"
            BevelOuter      =   0
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel lblBuyUnit 
            Height          =   300
            Left            =   5040
            TabIndex        =   60
            Top             =   3780
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "(일)"
            BevelOuter      =   0
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   615
         Left            =   30
         TabIndex        =   42
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
         Caption         =   " ▒ 물품기초정보"
         BevelOuter      =   1
         BevelInner      =   2
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   570
         Left            =   30
         TabIndex        =   43
         Top             =   7065
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
            TabIndex        =   44
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
            TabIndex        =   45
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
            TabIndex        =   46
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
            TabIndex        =   47
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
         TabIndex        =   48
         Top             =   30
         Width           =   8280
         _ExtentX        =   14605
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
         Caption         =   " ▒ 물품등록현황"
         BevelOuter      =   1
         BevelInner      =   2
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox cboKindF 
            Height          =   300
            Left            =   4200
            Style           =   2  '드롭다운 목록
            TabIndex        =   54
            Top             =   180
            Width           =   3765
         End
         Begin Threed.SSPanel SSPanel25 
            Height          =   300
            Left            =   3300
            TabIndex        =   55
            Top             =   180
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   529
            _Version        =   262144
            Caption         =   "물품분류"
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frm물품기초자료"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cStk As clsMstStk

Private Sub psListRefresh()
Dim sRow As Long

    gSql = "select a.stkcd, a.stknm, a.maker, b.custnm from mstSTK a, mstCUST b where a.custcd *= b.custcd"
    If Len(txtFind.Text) > 0 Then
        gSql = gSql & " and a.stknm like '%" & Trim(txtFind.Text) & "%'"
    End If
    If cboKindF.ListIndex > 0 Then
        gSql = gSql & " and a.kindcd = " & cboKindF.ItemData(cboKindF.ListIndex)
    End If
    gSql = gSql & " order by a.stknm"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                Call gsSpreadClear(spList, .RecordCount, True)
                While (Not .EOF)
                    sRow = sRow + 1
                    
                    spList.SetText 1, sRow, "" & .Fields("stkcd").Value
                    spList.SetText 2, sRow, "" & .Fields("stknm").Value
                    spList.SetText 3, sRow, "" & .Fields("maker").Value
                    spList.SetText 4, sRow, "" & .Fields("custnm").Value
                    
                    .MoveNext
                Wend
            Else
                Call gsSpreadClear(spList, 1, True)
            End If
            .Close
        End If
    End With

End Sub

Private Sub cboCust_Click()

    If cboCust.ListIndex >= 0 Then
        lblCustNm.Caption = cboCust.SubItem(cboCust.ListIndex, 1)
    End If

End Sub

Private Sub cboKind_Click()

    If cboKind.ListIndex >= 0 Then
        lblKindNm.Caption = cboKind.SubItem(cboKind.ListIndex, 1)
    End If

End Sub

Private Sub cboKindF_Click()

    MousePointer = vbHourglass
    Call psListRefresh
    MousePointer = vbDefault

End Sub

Private Sub cmdClear_Click()

    Call gsSetCustComboPV(cboCust, False)
    Call gsSetKindComboPV(cboKind, False)
    Call gsSetKindCombo(cboKindF, False)

    Call gsFieldClear(Me)
    cboRmd.ListIndex = -1
    cboBuyType.ListIndex = -1
    numRateQty.Text = 1
    
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
        If MsgBox("선택하신 물품자료를 삭제하시겠습니까 ?", vbYesNo + vbQuestion) = vbYes Then
            If cStk.cfDelete(Val(lblNo.Caption)) Then
                Call cmdClear_Click
                Call psListRefresh
            End If
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
        MsgBox "품명을 입력하세요.!", vbCritical
        Exit Sub
    End If
    
    MousePointer = vbHourglass
    With cStk
        .barcode = Trim(txtBarcode.Text)
        .buyamt = Val(numBuyAmt.Value)
        .buyday = Val(numBuyDay.Text)
        .buyioqty = Val(numRateQty.Text)
        .buytype = cboBuyType.ListIndex
        .buyunit = Trim(txtBuyUnit.Text)
        .custcd = Val(cboCust.Text)
        .delfg = chkDelete.Value
        .iounit = Trim(txtIoUnit.Text)
        .kindcd = Val(cboKind.Text)
        .maker = Trim(txtMaker.Text)
        .minbuyqty = Val(numMinQty.Text)
        .rmdfg = cboRmd.ListIndex
        .safeqty = Val(numSafeQty.Text)
        .stdamt = Val(numStdAmt.Value)
        .stkcd = Val(lblNo.Caption)
        .stknm = Trim(txtNm.Text)
        .stkspec = Trim(txtSpec.Text)
        
        If .cfSave Then
            lblNo.Caption = .stkcd
            cmdDelete.Enabled = True
            
            Call psListRefresh
        End If
    End With
    MousePointer = vbDefault
    Exit Sub
    
ErrcmdSave:
    MousePointer = vbDefault
    MsgBox Err.Description, vbCritical

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    Call gsEnterEsc_KeyPress(Me, KeyAscii, Me.Count)

End Sub

Private Sub Form_Load()
Dim sCnt As Integer

    Set cStk = New clsMstStk

    Me.KeyPreview = True
    Me.Show
    
    cboBuyType.Clear
    For sCnt = 0 To UBound(gBuyType)
        cboBuyType.AddItem gBuyType(sCnt)
    Next sCnt
    
    cboRmd.Clear
    For sCnt = 0 To UBound(gRmdType)
        cboRmd.AddItem gRmdType(sCnt)
    Next sCnt
    
    Call cmdClear_Click
    Call psListRefresh
    
End Sub

Private Sub spList_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim sData As Variant

    If Row > 0 And Col > 0 Then
        spList.GetText 1, Row, sData
        
        gSql = "select a.*, b.custnm, c.kindnm from mstSTK a, mstCUST b, mstSTKG c where a.stkcd = " & Val(sData) & _
               "   and a.custcd *= b.custcd and a.kindcd *= c.kindcd"
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                If Not .EOF Then
                    txtBarcode.Text = "" & .Fields("barcode").Value
                    numBuyAmt.Value = Val("" & .Fields("buyamt").Value)
                    numBuyDay.Text = Val("" & .Fields("buyday").Value)
                    numRateQty.Text = Val("" & .Fields("buyioqty").Value)
                    cboBuyType.ListIndex = Val("" & .Fields("buytype").Value)
                    txtBuyUnit.Text = "" & .Fields("buyunit").Value
                    
                    If Val("" & .Fields("custcd").Value) = 0 Then
                        cboCust.Text = ""
                        lblCustNm.Caption = ""
                    Else
                        cboCust.Text = Val("" & .Fields("custcd").Value)
                        lblCustNm.Caption = "" & .Fields("custnm").Value
                    End If
                    chkDelete.Value = Val("" & .Fields("delfg").Value)
                    txtIoUnit.Text = "" & .Fields("iounit").Value
                    If Val("" & .Fields("kindcd").Value) = 0 Then
                        cboKind.Text = ""
                        lblKindNm.Caption = ""
                    Else
                        cboKind.Text = Val("" & .Fields("kindcd").Value)
                        lblKindNm.Caption = "" & .Fields("kindnm").Value
                    End If
                    txtMaker.Text = "" & .Fields("maker").Value
                    numMinQty.Text = Val("" & .Fields("minbuyqty").Value)
                    cboRmd.ListIndex = Val("" & .Fields("rmdfg").Value)
                    numSafeQty.Text = Val("" & .Fields("safeqty").Value)
                    numStdAmt.Value = Val("" & .Fields("stdamt").Value)
                    lblNo.Caption = Val("" & .Fields("stkcd").Value)
                    txtNm.Text = "" & .Fields("stknm").Value
                    txtSpec.Text = "" & .Fields("stkspec").Value
                    lblWrtdt.Caption = "" & .Fields("wrtdt").Value
                    lblModdt.Caption = "" & .Fields("moddt").Value
                    
                    cmdSave.Enabled = True
                    cmdDelete.Enabled = True
                End If
                .Close
            End If
        End With
    End If
    
End Sub

Private Sub txtBuyUnit_Change()

    lblBuyUnit.Caption = "(" & txtBuyUnit.Text & ")"

End Sub

Private Sub txtNm_LostFocus()

    cmdSave.Enabled = Len(txtNm.Text)
    
End Sub
