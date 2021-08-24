VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{9255B445-567E-4A7A-9DCD-987EFAE369A8}#2.0#0"; "XCheckbutton.ocx"
Object = "{B88C4DC8-B707-435E-8B13-08058839823E}#2.0#0"; "XLabel.ocx"
Object = "{94265C54-C72D-40E9-95C4-FBB6BF532C26}#2.0#0"; "XGroupBox.ocx"
Object = "{1A3A9E7F-34C1-4F5C-BD80-63FA100EC4A0}#2.0#0"; "XCombobox.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form PIS306 
   BackColor       =   &H00FFFFFF&
   Caption         =   "재고현황"
   ClientHeight    =   9765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9765
   ScaleWidth      =   14880
   WindowState     =   2  '최대화
   Begin XLibrary_XGroupBox.XGroupBox grpMain 
      Height          =   9675
      Left            =   30
      Top             =   30
      Width           =   14790
      _ExtentX        =   26088
      _ExtentY        =   17066
      BackColor       =   16777215
      BorderColor     =   10070188
      BorderRoundNum  =   0
      BorderStyle     =   1
      TextColor       =   0
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      TextPosition    =   0
      TextCustomMargin=   4
      GroupBoxStyle   =   0
      TextBarColor1   =   12757903
      TextBarStyle    =   3
      TextBarColor2   =   11767328
      TextBarSymbol   =   0   'False
      TextBarSymbolColor=   16777215
      TextBarHeightMargin=   10
      MouseCursor     =   0
      TextBarMouseCursor=   0
      IconandTextMargin=   4
      BodyColor       =   16777215
      Enabled         =   -1  'True
      Begin BHButton.BHImageButton cmdFind 
         Height          =   375
         Left            =   11040
         TabIndex        =   8
         Top             =   810
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "조 회"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS306.frx":0000
         BackColor       =   12632319
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdExcel 
         Height          =   375
         Left            =   12270
         TabIndex        =   7
         Top             =   810
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "Excel"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS306.frx":17C2
         BackColor       =   12632319
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClose 
         Height          =   375
         Left            =   13500
         TabIndex        =   6
         Top             =   810
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "닫 기"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS306.frx":2F84
         BackColor       =   12632319
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClear 
         Height          =   375
         Left            =   9810
         TabIndex        =   5
         Top             =   810
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "화면지움"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS306.frx":4746
         BackColor       =   12632319
         ImgOutLineSize  =   3
      End
      Begin FPSpreadADO.fpSpread spList 
         CausesValidation=   0   'False
         Height          =   8355
         Left            =   90
         TabIndex        =   1
         Tag             =   "20001"
         Top             =   1260
         Width           =   14640
         _Version        =   524288
         _ExtentX        =   25823
         _ExtentY        =   14737
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   9
         MaxRows         =   489
         Protect         =   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "PIS306.frx":5F08
         VisibleCols     =   3
         VisibleRows     =   10
         CellNoteIndicatorColor=   16576
      End
      Begin XLibrary_XGroupBox.XGroupBox grpFind 
         Height          =   645
         Left            =   90
         Top             =   90
         Width           =   14640
         _ExtentX        =   25823
         _ExtentY        =   1138
         BackColor       =   16311512
         BorderColor     =   10070188
         BorderRoundNum  =   0
         BorderStyle     =   1
         TextColor       =   0
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         TextPosition    =   0
         TextCustomMargin=   4
         GroupBoxStyle   =   0
         TextBarColor1   =   12757903
         TextBarStyle    =   3
         TextBarColor2   =   11767328
         TextBarSymbol   =   0   'False
         TextBarSymbolColor=   16777215
         TextBarHeightMargin=   10
         MouseCursor     =   0
         TextBarMouseCursor=   0
         IconandTextMargin=   4
         BodyColor       =   16777215
         Enabled         =   -1  'True
         Begin TDBDate6Ctl.TDBDate dtpTodt 
            Height          =   315
            Left            =   2820
            TabIndex        =   10
            Top             =   180
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   556
            Calendar        =   "PIS306.frx":66C2
            Caption         =   "PIS306.frx":67A9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "PIS306.frx":680C
            Keys            =   "PIS306.frx":682A
            Spin            =   "PIS306.frx":6888
            AlignHorizontal =   0
            AlignVertical   =   2
            Appearance      =   2
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "yyyy-mm-dd"
            EditMode        =   1
            Enabled         =   -1
            ErrorBeep       =   0
            FirstMonth      =   4
            ForeColor       =   -2147483640
            Format          =   "yyyy-mm-dd"
            HighlightText   =   0
            IMEMode         =   3
            MarginBottom    =   1
            MarginLeft      =   5
            MarginRight     =   5
            MarginTop       =   1
            MaxDate         =   2958465
            MinDate         =   -657434
            MousePointer    =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            PromptChar      =   "_"
            ReadOnly        =   0
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "2015-07-15"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   42200
            CenturyMode     =   2
         End
         Begin TDBDate6Ctl.TDBDate dtpFrdt 
            Height          =   315
            Left            =   1380
            TabIndex        =   9
            Top             =   180
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   556
            Calendar        =   "PIS306.frx":68B0
            Caption         =   "PIS306.frx":6997
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "PIS306.frx":69FA
            Keys            =   "PIS306.frx":6A18
            Spin            =   "PIS306.frx":6A76
            AlignHorizontal =   0
            AlignVertical   =   2
            Appearance      =   2
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "yyyy-mm-dd"
            EditMode        =   1
            Enabled         =   -1
            ErrorBeep       =   0
            FirstMonth      =   4
            ForeColor       =   -2147483640
            Format          =   "yyyy-mm-dd"
            HighlightText   =   0
            IMEMode         =   3
            MarginBottom    =   1
            MarginLeft      =   5
            MarginRight     =   5
            MarginTop       =   1
            MaxDate         =   2958465
            MinDate         =   -657434
            MousePointer    =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            PromptChar      =   "_"
            ReadOnly        =   0
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "2015-07-15"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   42200
            CenturyMode     =   2
         End
         Begin XLibrary_XCheckButton.XCheckButton chkAll 
            Height          =   315
            Left            =   7740
            TabIndex        =   4
            Top             =   180
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16311512
            CbBackColor1    =   15132390
            CbBorderColor   =   8409372
            CbBorderStyle   =   0
            Text            =   "전체품목"
            TextColor       =   0
            CbTextMargin    =   4
            CbBackStyle     =   1
            CbGDirection    =   0
            CbBackColor2    =   16777215
            CheckColor      =   2203937
            CheckCustomColor=   2998317
            Value           =   -1  'True
            CbOverEffect    =   -1  'True
            CbOverEffectGDtn=   0
            CbOverColor1    =   10280958
            CbOverColor2    =   3388664
            MouseCursor     =   0
            ToolTipOpacity  =   100
            ToolTipIcon     =   1
            ToolTipPopupTime=   -1
            ToolTipHoverTime=   -1
            ToolTipBackColor=   14811135
            ToolTipForeColor=   0
            ToolTipStyle    =   2
            ToolTipCentered =   -1  'True
            ToolTipTitleText=   ""
            ToolTipBodyText =   ""
            Enabled         =   -1  'True
            EnabledAutoStyle=   -1  'True
            EnCbBackColor   =   14215660
            EnCbBorderColor =   10070188
            EnCheckColor    =   10070188
            EnTextColor     =   10070188
            CheckStyle      =   0
            ControlType     =   0
            AutoSize        =   0   'False
         End
         Begin XLibrary_XLabel.XLabel XLabel11 
            Height          =   315
            Left            =   4440
            TabIndex        =   3
            Top             =   180
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "재고구분"
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TextColor       =   0
            IconAndTextMargin=   8
            TextAlign       =   0
            TextAlignMargin =   0
            Focus           =   0   'False
            MouseCursor     =   0
            ToolTipIcon     =   0
            ToolTipPopupTime=   -1
            ToolTipHoverTime=   -1
            ToolTipBackColor=   14811135
            ToolTipForeColor=   0
            ToolTipOpacity  =   100
            ToolTipStyle    =   0
            ToolTipCentered =   0   'False
            ToolTipTitleText=   "Title"
            ToolTipBodyText =   ""
            TextBackColor1  =   -2147483629
            TextBackColor2  =   -2147483629
            TextBackMargin  =   4
            TextBackStyle   =   0
            Enabled         =   -1  'True
         End
         Begin XLibrary_XComboBox.XComboBox cboRmd 
            Height          =   315
            Left            =   5850
            TabIndex        =   2
            Top             =   180
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            BackColor       =   16777215
            BorderColor     =   16744576
            BtnBackColor1   =   16777215
            BtnBackStyle    =   3
            Text            =   ""
            BtnBorderColor  =   12632256
            BtnBorderStyle  =   1
            BtnBackColor2   =   15000804
            BtnSymbolColor  =   8388608
            BtnSymbolStyle  =   2
            UpListShow      =   0   'False
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowItemNum     =   5
            AutoSel         =   0   'False
            TextEdit        =   0   'False
            BtnMouseCursor  =   2
            ToolTipIcon     =   1
            ToolTipPopupTime=   -1
            ToolTipHoverTime=   800
            ToolTipBackColor=   16777215
            ToolTipForeColor=   0
            ToolTipOpacity  =   100
            ToolTipStyle    =   2
            ToolTipCentered =   0   'False
            ToolTipTitleText=   ""
            ToolTipBodyText =   ""
            TextColor       =   0
            ListBgColor     =   16777215
            ListTextColor   =   0
            Enabled         =   -1  'True
         End
         Begin XLibrary_XLabel.XLabel XLabel6 
            Height          =   315
            Left            =   210
            TabIndex        =   0
            Top             =   180
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "수불기간"
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TextColor       =   0
            IconAndTextMargin=   8
            TextAlign       =   0
            TextAlignMargin =   0
            Focus           =   0   'False
            MouseCursor     =   0
            ToolTipIcon     =   0
            ToolTipPopupTime=   -1
            ToolTipHoverTime=   -1
            ToolTipBackColor=   14811135
            ToolTipForeColor=   0
            ToolTipOpacity  =   100
            ToolTipStyle    =   0
            ToolTipCentered =   0   'False
            ToolTipTitleText=   "Title"
            ToolTipBodyText =   ""
            TextBackColor1  =   -2147483629
            TextBackColor2  =   -2147483629
            TextBackMargin  =   4
            TextBackStyle   =   0
            Enabled         =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "PIS306"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()

    dtpFrdt.Value = Format(gfSystemDate, "yyyy-MM") & "-01"
    dtpTodt.Value = gfSystemDate
    
    Call gsSpreadClear(spList, 0, True)
    Call gsButtonEnable(cmdFind, True)
    Call gsButtonEnable(cmdExcel, False)
    
    grpFind.Enabled = True
    
End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdExcel_Click()

    Call gsSpreadToExcel(spList, Me.Caption)

End Sub

Private Sub cmdFind_Click()
Dim sRow As Long, sRmdQty As Double, sUnitQty As Double, sYear As String, sFrDate As String, sToDate As String, sSafeQty As Double

    MousePointer = vbHourglass
    sYear = Format(dtpFrdt.Value, "yyyy")
    sFrDate = Format(dtpFrdt.Value, "yyyyMMdd")
    sToDate = Format(dtpTodt.Value, "yyyyMMdd")
    
    gSql = "SELECT X.CD_ITEM AS STKCD,A.PREVQTY,A.ENTQTY,A.OUTQTY,X.NM_ITEM AS STKNM,X.UNIT_SO AS UNIT, X.UNIT_SO_FACT AS UNITRATE,                 " & vbNewLine & _
           "       X.QT_SSTOCK AS SAFEQTY FROM (                                                                                                    " & vbNewLine & _
           "SELECT Y.STKCD, SUM(Y.PREVQTY) AS PREVQTY, SUM(Y.ENTQTY) AS ENTQTY, SUM(Y.OUTQTY) AS OUTQTY                                             " & vbNewLine & _
           "  FROM (                                                                                                                                " & vbNewLine & _
           "    SELECT Z.STKCD, SUM(Z.PREVQTY+Z.ENTQTY-Z.OUTQTY) AS PREVQTY, 0 AS ENTQTY, 0 AS OUTQTY                                               " & vbNewLine & _
           "      FROM (                                                                                                                            " & vbNewLine & _
           "        SELECT STKCD, PREVQTY, 0 AS ENTQTY, 0 AS OUTQTY FROM S2PIS409 WHERE RMDYEAR='" & sYear & "'                                     " & vbNewLine & _
           "        UNION ALL                                                                                                                       " & vbNewLine
    If cboRmd.ListIndex = 0 Then
        gSql = gSql & _
               "        SELECT STKCD, 0 AS PREVQTY, SUM(IQTY_SO) AS ENTQTY, 0 AS OUTQTY FROM S2PIS201                                               " & vbNewLine & _
               "         WHERE SUBSTR(ENTDT,1,4)='" & sYear & "' AND ENTDT<'" & sFrDate & "' GROUP BY STKCD                                         " & vbNewLine
    Else
        gSql = gSql & _
               "        SELECT STKCD, 0 AS PREVQTY, SUM(ENTQTY) AS ENTQTY, 0 AS OUTQTY FROM S2PIS401                                                " & vbNewLine & _
               "         WHERE SUBSTR(CHULDT,1,4)='" & sYear & "' AND CHULDT<'" & sFrDate & "' GROUP BY STKCD                                       " & vbNewLine
    End If
    gSql = gSql & _
           "        UNION ALL                                                                                                                       " & vbNewLine & _
           "        SELECT STKCD, 0 AS PREVQTY, 0 AS ENTQTY,                                                                                        " & vbNewLine & _
           "               SUM(NVL(TESTQTY,0)+NVL(FREEQTY,0)+NVL(QCQTY,0)+NVL(RETESTQTY,0)+NVL(MANUQTY,0)+NVL(MACHQTY,0)+NVL(HANDQTY,0)) AS OUTQTY  " & vbNewLine & _
           "          FROM S2PIS313 WHERE SUBSTR(WORKDT,1,4)='" & sYear & "' AND WORKDT<'" & sFrDate & "' GROUP BY STKCD                            " & vbNewLine & _
           "    ) Z GROUP BY Z.STKCD                                                                                                                " & vbNewLine & _
           "    UNION ALL                                                                                                                           " & vbNewLine
    If cboRmd.ListIndex = 0 Then
        gSql = gSql & _
               "    SELECT STKCD, 0 AS PREVQTY, SUM(IQTY_SO) AS ENTQTY, 0 AS OUTQTY                                                                 " & vbNewLine & _
               "      FROM S2PIS201 WHERE ENTDT BETWEEN '" & sFrDate & "' AND '" & sToDate & "' GROUP BY STKCD                                      " & vbNewLine
    Else
        gSql = gSql & _
               "    SELECT STKCD, 0 AS PREVQTY, SUM(ENTQTY) AS ENTQTY, 0 AS OUTQTY                                                                  " & vbNewLine & _
               "      FROM S2PIS401 WHERE CHULDT BETWEEN '" & sFrDate & "' AND '" & sToDate & "' GROUP BY STKCD                                     " & vbNewLine
    End If
    gSql = gSql & _
           "    UNION ALL                                                                                                                           " & vbNewLine & _
           "    SELECT STKCD, 0 AS PREVQTY, 0 AS ENTQTY,                                                                                            " & vbNewLine & _
           "           SUM(NVL(TESTQTY,0)+NVL(FREEQTY,0)+NVL(QCQTY,0)+NVL(RETESTQTY,0)+NVL(MANUQTY,0)+NVL(MACHQTY,0)+NVL(HANDQTY,0)) AS OUTQTY      " & vbNewLine & _
           "      FROM S2PIS313 WHERE WORKDT BETWEEN '" & sFrDate & "' AND '" & sToDate & "' GROUP BY STKCD                                         " & vbNewLine & _
           ") Y GROUP BY Y.STKCD                                                                                                                    " & vbNewLine & _
           ") A " & IIf(chkAll.Value, "RIGHT", "") & " JOIN " & gTBLstk & " X ON A.STKCD=X.CD_ITEM " & gERPStkCondition & vbNewLine & _
           " ORDER BY X.CD_ITEM"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                gPrgBar.Max = .RecordCount:     gPrgBar.Value = 0
                gPrgBar.Visible = True:         gPrgBar.Refresh
                Call gsSpreadClear(spList, .RecordCount, True)
                While (Not .EOF)
                    sRow = sRow + 1
                    
                    spList.SetText 1, sRow, "" & .Fields("STKCD").Value
                    spList.SetText 2, sRow, "" & .Fields("STKNM").Value
                    spList.SetText 3, sRow, "" & .Fields("UNIT").Value
                    spList.SetText 4, sRow, gfQtyOutputStr(Val("" & .Fields("PREVQTY").Value))
                    spList.SetText 5, sRow, gfQtyOutputStr(Val("" & .Fields("ENTQTY").Value))
                    spList.SetText 6, sRow, gfQtyOutputStr(Val("" & .Fields("OUTQTY").Value))
                    
                    sRmdQty = Val("" & .Fields("PREVQTY").Value) + Val("" & .Fields("ENTQTY").Value) - Val("" & .Fields("OUTQTY").Value)
                    spList.SetText 7, sRow, gfQtyOutputStr(sRmdQty)
                    
                    If Val("" & .Fields("UNITRATE").Value) <> 0 Then
                        sUnitQty = sRmdQty / Val("" & .Fields("UNITRATE").Value)
                        sSafeQty = Val("" & .Fields("SAFEQTY").Value) * Val("" & .Fields("UNITRATE").Value)
                    Else
                        sUnitQty = sRmdQty
                        sSafeQty = Val("" & .Fields("SAFEQTY").Value)
                    End If
                    spList.SetText 8, sRow, gfQtyInputStr(sUnitQty)
                    spList.SetText 9, sRow, gfQtyOutputStr(sSafeQty)
                    
                    .MoveNext
                Wend
                gPrgBar.Visible = False
                
                Call gsButtonEnable(cmdFind, False)
                Call gsButtonEnable(cmdExcel, True)
                grpFind.Enabled = False
            Else
                Call gsSpreadClear(spList, 0, True)
                MsgBox "조건에 해당하는 입고자료가 없습니다.!", vbCritical
            End If
            .Close
        End If
    End With
    MousePointer = vbDefault

End Sub

Private Sub Form_Activate()

    frmMain.lblTitle.Text = Me.Caption

End Sub

Private Sub Form_Load()
Dim sCol As Integer

    grpMain.BorderColor = gGrpLineColor
    grpMain.BackColor = gGrpBackColor

    Me.Show
    
    With spList
        .UserColAction = UserColActionSort
        For sCol = 1 To .MaxCols
            .ColUserSortIndicator(sCol) = ColUserSortIndicatorAscending
        Next sCol
    End With
    
    With cboRmd
        .Clear
        .AddItem "전체재고", 0
        .AddItem "현장재고", 1
        
        .ListIndex = 0
    End With
    
    With spList
        .SetText 7, SpreadHeader, "재고량"
        .SetText 8, SpreadHeader, "재고량" & vbNewLine & "(재고단위)"
        .SetText 9, SpreadHeader, "적정재고" & vbNewLine & "(재고단위)"
    End With
    
    chkAll.Value = False
    
    Call cmdClear_Click
    
End Sub

Private Sub Form_Resize()
On Error Resume Next

    grpMain.Left = (Me.ScaleWidth - grpMain.Width) / 2
    grpMain.Height = Me.ScaleHeight - 50
    spList.Height = (grpMain.Height - spList.Top) - 50
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmMain.lblTitle.Text = ""
    
End Sub

Private Sub spList_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim sGetVal As Variant

    If Row > 0 And Col > 0 Then
        With PIS305
            .dtpFrdt.Value = dtpFrdt.Value
            .dtpTodt.Value = dtpTodt.Value
            spList.GetText 1, Row, sGetVal
            .txtStkcd.Text = Trim(sGetVal)
            spList.GetText 2, Row, sGetVal
            .txtStkNm.Text = Trim(sGetVal)
            
            Call .fsFindCall
            .ZOrder 0
        End With
    End If

End Sub
