VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{BD146989-F30B-4134-B202-680CC90638EF}#2.0#0"; "XTextBox.ocx"
Object = "{9255B445-567E-4A7A-9DCD-987EFAE369A8}#2.0#0"; "XCheckbutton.ocx"
Object = "{B88C4DC8-B707-435E-8B13-08058839823E}#2.0#0"; "XLabel.ocx"
Object = "{94265C54-C72D-40E9-95C4-FBB6BF532C26}#2.0#0"; "XGroupBox.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form PIS309 
   BackColor       =   &H00FFFFFF&
   Caption         =   "창고불출현황"
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
      Begin BHButton.BHImageButton cmdDelete 
         Height          =   375
         Left            =   12300
         TabIndex        =   10
         Top             =   810
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "삭 제"
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
         TransparentPicture=   "PIS309.frx":0000
         BackColor       =   12632319
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdFind 
         Height          =   375
         Left            =   9840
         TabIndex        =   9
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
         TransparentPicture=   "PIS309.frx":17C2
         BackColor       =   12632319
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdExcel 
         Height          =   375
         Left            =   11070
         TabIndex        =   8
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
         TransparentPicture=   "PIS309.frx":2F84
         BackColor       =   12632319
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClose 
         Height          =   375
         Left            =   13530
         TabIndex        =   7
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
         TransparentPicture=   "PIS309.frx":4746
         BackColor       =   12632319
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClear 
         Height          =   375
         Left            =   8610
         TabIndex        =   6
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
         TransparentPicture=   "PIS309.frx":5F08
         BackColor       =   12632319
         ImgOutLineSize  =   3
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
            Left            =   2850
            TabIndex        =   12
            Top             =   180
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   556
            Calendar        =   "PIS309.frx":76CA
            Caption         =   "PIS309.frx":77B1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "PIS309.frx":7814
            Keys            =   "PIS309.frx":7832
            Spin            =   "PIS309.frx":7890
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
            Left            =   1410
            TabIndex        =   13
            Top             =   180
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   556
            Calendar        =   "PIS309.frx":78B8
            Caption         =   "PIS309.frx":799F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "PIS309.frx":7A02
            Keys            =   "PIS309.frx":7A20
            Spin            =   "PIS309.frx":7A7E
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
         Begin BHButton.BHImageButton cmdStk 
            Height          =   315
            Left            =   6960
            TabIndex        =   11
            Top             =   180
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            Caption         =   "..."
            CaptionChecked  =   "BHImageButton1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TransparentPicture=   "PIS309.frx":7AA6
            BackColor       =   14737632
            AlphaColor      =   16777215
            ImgOutLineSize  =   3
         End
         Begin XLibrary_XCheckButton.XCheckButton chkRmd 
            Height          =   315
            Left            =   10470
            TabIndex        =   5
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
            Text            =   "재고(有)"
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
         Begin XLibrary_XTextBox.XTextBox txtStkNm 
            Height          =   315
            Left            =   7320
            TabIndex        =   4
            Top             =   180
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            BackColor       =   14737632
            BorderColor     =   16744576
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            BorderTextMargin=   4
            PasswordChar    =   ""
            MaxLength       =   0
            MouseCursor     =   4
            TextColor       =   0
            ToolTipOpacity  =   100
            ToolTipIcon     =   2
            ToolTipPopupTime=   -1
            ToolTipHoverTime=   -1
            ToolTipBackColor=   16777215
            ToolTipForeColor=   0
            ToolTipStyle    =   3
            ToolTipCentered =   0   'False
            ToolTipTitleText=   ""
            ToolTipBodyText =   ""
            Locked          =   -1  'True
            Mask            =   0
            PromptChar      =   "_"
            WrongSound      =   0
            CustomSound     =   ""
            MaskShow        =   0   'False
            MaskColor       =   33023
            CustomMask      =   ""
            TextAlign       =   0
            Enabled         =   -1  'True
         End
         Begin XLibrary_XLabel.XLabel XLabel4 
            Height          =   315
            Left            =   4470
            TabIndex        =   3
            Top             =   180
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "품목코드"
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
         Begin XLibrary_XTextBox.XTextBox txtStkcd 
            Height          =   315
            Left            =   5880
            TabIndex        =   2
            Top             =   180
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            BackColor       =   16777215
            BorderColor     =   16744576
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            BorderTextMargin=   4
            PasswordChar    =   ""
            MaxLength       =   0
            MouseCursor     =   4
            TextColor       =   0
            ToolTipOpacity  =   100
            ToolTipIcon     =   2
            ToolTipPopupTime=   -1
            ToolTipHoverTime=   -1
            ToolTipBackColor=   16777215
            ToolTipForeColor=   0
            ToolTipStyle    =   3
            ToolTipCentered =   0   'False
            ToolTipTitleText=   ""
            ToolTipBodyText =   ""
            Locked          =   0   'False
            Mask            =   0
            PromptChar      =   "_"
            WrongSound      =   0
            CustomSound     =   ""
            MaskShow        =   0   'False
            MaskColor       =   33023
            CustomMask      =   ""
            TextAlign       =   0
            Enabled         =   -1  'True
         End
         Begin XLibrary_XLabel.XLabel XLabel6 
            Height          =   315
            Left            =   210
            TabIndex        =   1
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
      Begin FPSpreadADO.fpSpread spList 
         CausesValidation=   0   'False
         Height          =   8355
         Left            =   90
         TabIndex        =   0
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
         MaxCols         =   13
         MaxRows         =   489
         Protect         =   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "PIS309.frx":9268
         VisibleCols     =   3
         VisibleRows     =   10
         CellNoteIndicatorColor=   16576
      End
   End
End
Attribute VB_Name = "PIS309"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub fsFindCall()

    Call cmdFind_Click

End Sub

Private Sub cmdClear_Click()

    dtpFrdt.Value = Format(gfSystemDate, "yyyy-MM") & "-01"
    dtpTodt.Value = gfSystemDate
    
    txtStkcd.Text = ""
    txtStkNm.Text = ""
    
    Call gsSpreadClear(spList, 0, True)
    Call gsButtonEnable(cmdExcel, False)
    Call gsButtonEnable(cmdDelete, False)
    
    grpFind.Enabled = True
    
End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdDelete_Click()
Dim sGetVal As Variant, cPis401 As clsPis401

    MousePointer = vbHourglass
    If MsgBox("선택하신 창고불출자료를 삭제하시겠습니까 ?", vbYesNo + vbQuestion) = vbYes Then
        Set cPis401 = New clsPis401
        spList.GetText 1, spList.ActiveRow, sGetVal:    cPis401.chuldt = Format(sGetVal, "yyyyMMdd")
        spList.GetText 11, spList.ActiveRow, sGetVal:   cPis401.chulseq = Val(sGetVal)
        If Len(cPis401.chuldt) > 0 And cPis401.chulseq > 0 Then
            spList.GetText 6, spList.ActiveRow, sGetVal:    cPis401.entqty = Val(Str(sGetVal))
            spList.GetText 12, spList.ActiveRow, sGetVal:   cPis401.entdt = Trim(sGetVal)
            spList.GetText 13, spList.ActiveRow, sGetVal:   cPis401.entseq = Val(sGetVal)
            
            Call cDb.csBegin
            If cPis401.cfDelete Then
                Call cDb.csCommit
                Call cmdFind_Click
            Else
                cDb.csRollback
            End If
        End If
    End If
    MousePointer = vbDefault
    
End Sub

Private Sub cmdExcel_Click()

    Call gsSpreadToExcel(spList, Me.Caption)

End Sub

Private Sub cmdFind_Click()
Dim sRow As Long, sRmdQty As Double, sFrDt As String, sToDt As String

    sFrDt = Format(dtpFrdt.Value, "yyyyMMdd")
    sToDt = Format(dtpTodt.Value, "yyyyMMdd")

    MousePointer = vbHourglass
    If gWorkArea Then
        gSql = "SELECT A.*, X.NM_ITEM AS STKNM, B.LOTNO, B.IQTY_SO AS INQTY, D.EMPNM FROM S2PIS401 A    " & vbNewLine & _
               "       LEFT JOIN S2PIS201 B ON A.ENTDT=B.ENTDT AND A.ENTSEQ=B.ENTSEQ                    " & vbNewLine & _
               "       LEFT JOIN " & gTBLstk & " X ON A.STKCD=X.CD_ITEM " & gERPStkCondition & "        " & vbNewLine & _
               "       LEFT JOIN S2COM006 D ON A.EMPID=D.EMPID                                          " & vbNewLine
    Else
        gSql = "SELECT A.*, X.NM_ITEM AS STKNM, B.LOTNO, B.IQTY_SO AS INQTY, D.USER_NM FROM S2PIS401 A  " & vbNewLine & _
               "       LEFT JOIN S2PIS201 B ON A.ENTDT=B.ENTDT AND A.ENTSEQ=B.ENTSEQ                    " & vbNewLine & _
               "       LEFT JOIN " & gTBLstk & " X ON A.STKCD=X.CD_ITEM " & gERPStkCondition & "        " & vbNewLine & _
               "       LEFT JOIN " & gKahpUserTable & " D ON A.EMPID=D.USERID                           " & vbNewLine
    End If
    gSql = gSql & " WHERE A.CHULDT BETWEEN '" & sFrDt & "' AND '" & sToDt & "'                          " & vbNewLine
    If Len(txtStkcd.Text) > 0 Then
        gSql = gSql & "  AND A.STKCD='" & Trim(txtStkcd.Text) & "'                                      " & vbNewLine
    End If
    If chkRmd.Value Then
        gSql = gSql & "  AND (A.USEQTY IS NULL OR A.ENTQTY>A.USEQTY)                                    " & vbNewLine
    End If
    gSql = gSql & " ORDER BY A.CHULDT,A.STKCD,A.CHULSEQ"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                gPrgBar.Max = .RecordCount:     gPrgBar.Value = 0
                gPrgBar.Visible = True:         gPrgBar.Refresh
                Call gsSpreadClear(spList, .RecordCount, True)
                
                While (Not .EOF)
                    sRow = sRow + 1
                    
                    spList.SetText 1, sRow, "" & Format(.Fields("CHULDT").Value, "####-##-##")
                    spList.SetText 2, sRow, "" & .Fields("STKCD").Value
                    spList.SetText 3, sRow, "" & .Fields("STKNM").Value
                    spList.SetText 4, sRow, "" & .Fields("LOTNO").Value
                    spList.SetText 5, sRow, gfQtyOutputStr(Val("" & .Fields("INQTY").Value))
                    spList.SetText 6, sRow, gfQtyOutputStr(Val("" & .Fields("ENTQTY").Value))
                    spList.SetText 7, sRow, gfQtyOutputStr(Val("" & .Fields("USEQTY").Value))
                    sRmdQty = Val("" & .Fields("ENTQTY").Value) - Val("" & .Fields("USEQTY").Value)
                    spList.SetText 8, sRow, gfQtyOutputStr(sRmdQty)
                    spList.SetText 9, sRow, "" & .Fields("EMPNM").Value
                    spList.SetText 10, sRow, "" & .Fields("MODDT").Value
                    spList.SetText 11, sRow, Val("" & .Fields("CHULSEQ").Value)
                    spList.SetText 12, sRow, "" & .Fields("ENTDT").Value
                    spList.SetText 13, sRow, Val("" & .Fields("ENTSEQ").Value)
                    
                    .MoveNext
                Wend
                gPrgBar.Visible = False
                Call gsButtonEnable(cmdExcel, True)
                grpFind.Enabled = False
            Else
                Call gsSpreadClear(spList, 0, True)
                MsgBox "조건에 해당하는 자료가 없습니다.!", vbCritical
            End If
            .Close
        End If
    End With
    MousePointer = vbDefault

End Sub

Private Sub cmdStk_Click()

    hlpStkList.Tag = "one"
    hlpStkList.Show vbModal
    
    If Len(gHelpCode) > 0 Then
        txtStkcd.Text = gHelpCode
        txtStkNm.Text = gfStkName(gHelpCode)
    Else
        txtStkcd.Text = ""
        txtStkNm.Text = ""
    End If

End Sub

Private Sub Form_Activate()

    frmMain.lblTitle.Text = Me.Caption

End Sub

Private Sub Form_Load()
Dim sCol As Integer

    grpMain.BorderColor = gGrpLineColor
    grpMain.BackColor = gGrpBackColor

    Me.Show
    
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

Private Sub spList_Click(ByVal Col As Long, ByVal Row As Long)
Dim sGetVal As Variant

    If Row > 0 And Col > 0 Then
        spList.GetText 7, spList.ActiveRow, sGetVal
        Call gsButtonEnable(cmdDelete, (Val(Str(sGetVal)) = 0))
    End If

End Sub

Private Sub txtStkcd_LostFocus()

    txtStkNm.Text = gfStkName(Trim(txtStkcd.Text))
    If Len(txtStkNm.Text) = 0 Then
        txtStkcd.Text = ""
    End If

End Sub


