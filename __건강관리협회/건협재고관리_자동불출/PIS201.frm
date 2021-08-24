VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{B88C4DC8-B707-435E-8B13-08058839823E}#2.0#0"; "XLabel.ocx"
Object = "{94265C54-C72D-40E9-95C4-FBB6BF532C26}#2.0#0"; "XGroupBox.ocx"
Object = "{1A3A9E7F-34C1-4F5C-BD80-63FA100EC4A0}#2.0#0"; "XCombobox.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form PIS201 
   BackColor       =   &H00FFFFFF&
   Caption         =   "입고자료등록"
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
      Begin BHButton.BHImageButton cmdClear 
         Height          =   375
         Left            =   8610
         TabIndex        =   8
         Top             =   840
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
         TransparentPicture=   "PIS201.frx":0000
         BackColor       =   12632319
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClose 
         Height          =   375
         Left            =   13530
         TabIndex        =   7
         Top             =   840
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
         TransparentPicture=   "PIS201.frx":17C2
         BackColor       =   12632319
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdCancel 
         Height          =   375
         Left            =   12300
         TabIndex        =   6
         Top             =   840
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "입고취소"
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
         TransparentPicture=   "PIS201.frx":2F84
         BackColor       =   12632319
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdEnter 
         Height          =   375
         Left            =   11070
         TabIndex        =   5
         Top             =   840
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "입고처리"
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
         TransparentPicture=   "PIS201.frx":4746
         BackColor       =   12632319
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdFind 
         Height          =   375
         Left            =   9840
         TabIndex        =   4
         Top             =   840
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
         TransparentPicture=   "PIS201.frx":5F08
         BackColor       =   12632319
         ImgOutLineSize  =   3
      End
      Begin FPSpreadADO.fpSpread spList 
         CausesValidation=   0   'False
         Height          =   8295
         Left            =   90
         TabIndex        =   1
         Tag             =   "20001"
         Top             =   1290
         Width           =   14640
         _Version        =   524288
         _ExtentX        =   25823
         _ExtentY        =   14631
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
         MaxCols         =   27
         MaxRows         =   489
         Protect         =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "PIS201.frx":76CA
         VisibleCols     =   3
         VisibleRows     =   10
         CellNoteIndicatorColor=   16576
      End
      Begin XLibrary_XGroupBox.XGroupBox XGroupBox3 
         Height          =   675
         Left            =   90
         Top             =   90
         Width           =   14640
         _ExtentX        =   25823
         _ExtentY        =   1191
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
            Left            =   2910
            TabIndex        =   10
            Top             =   180
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   556
            Calendar        =   "PIS201.frx":8779
            Caption         =   "PIS201.frx":8860
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "PIS201.frx":88C3
            Keys            =   "PIS201.frx":88E1
            Spin            =   "PIS201.frx":893F
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
            Left            =   1500
            TabIndex        =   9
            Top             =   180
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   556
            Calendar        =   "PIS201.frx":8967
            Caption         =   "PIS201.frx":8A4E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "PIS201.frx":8AB1
            Keys            =   "PIS201.frx":8ACF
            Spin            =   "PIS201.frx":8B2D
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
         Begin XLibrary_XComboBox.XComboBox cboProc 
            Height          =   315
            Left            =   5820
            TabIndex        =   3
            Top             =   180
            Width           =   1545
            _ExtentX        =   2725
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
         Begin XLibrary_XLabel.XLabel XLabel1 
            Height          =   315
            Left            =   4410
            TabIndex        =   2
            Top             =   180
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "처리구분"
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
         Begin XLibrary_XLabel.XLabel XLabel6 
            Height          =   315
            Left            =   210
            TabIndex        =   0
            Top             =   180
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "입고기간"
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
Attribute VB_Name = "PIS201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Dim cPis201 As clsPis201, sRow As Long, sGetVal As Variant, sReturn As Boolean

    MousePointer = vbHourglass
    If MsgBox("선택하신 자료를 입고취소 처리하시겠습니까 ?", vbYesNo + vbQuestion) = vbYes Then
        Set cPis201 = New clsPis201
        
        With spList
            For sRow = 1 To .MaxRows
                .GetText 1, sRow, sGetVal
                If Val(sGetVal) > 0 Then
                    .GetText 12, sRow, sGetVal
                    If Len(sGetVal) > 0 Then
                        .GetText 13, sRow, sGetVal:     cPis201.CD_COMPANY = Trim(sGetVal)
                        .GetText 14, sRow, sGetVal:     cPis201.CD_PLANT = Trim(sGetVal)
                        .GetText 15, sRow, sGetVal:     cPis201.CD_BIZAREA = Trim(sGetVal)
                        .GetText 16, sRow, sGetVal:     cPis201.NO_IO = Trim(sGetVal)
                        .GetText 17, sRow, sGetVal:     cPis201.NO_IOLINE = Trim(sGetVal)
                        .GetText 27, sRow, sGetVal:     cPis201.NO_IOLINE2 = Val(sGetVal)
                        
                        Call cDb.csBegin
                        sReturn = cPis201.cfDelete
                        If sReturn Then
                            Call cDb.csCommit
                            
                            .SetText 1, sRow, ""
                            .SetText 12, sRow, ""
                        Else
                            Call cDb.csRollback
                            Exit For
                        End If
                    End If
                End If
            Next sRow
        End With
        
        If sReturn Then Call cmdFind_Click
    End If
    MousePointer = vbDefault
    
End Sub

Private Sub cmdClear_Click()

    Call gsSpreadClear(spList, 0, True)
    
    Call gsButtonEnable(cmdFind, True)
    Call gsButtonEnable(cmdEnter, False)
    Call gsButtonEnable(cmdCancel, False)
    
End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdEnter_Click()
Dim cPis201 As clsPis201, sRow As Long, sGetVal As Variant, sRateQty As Single, sReturn As Boolean

    MousePointer = vbHourglass
    If MsgBox("선택하신 자료를 입고처리하시겠습니까 ?", vbYesNo + vbQuestion) = vbYes Then
        Set cPis201 = New clsPis201
        
        With spList
            For sRow = 1 To .MaxRows
                .GetText 1, sRow, sGetVal
                If Val(sGetVal) > 0 Then
                    .GetText 12, sRow, sGetVal
                    If Len(sGetVal) = 0 Then
                        .GetText 2, sRow, sGetVal:      cPis201.entdt = Format(sGetVal, "yyyyMMdd")
                        cPis201.entseq = 0
                        .GetText 3, sRow, sGetVal:      cPis201.cstcd = Trim(sGetVal)
                        .GetText 4, sRow, sGetVal:      cPis201.cstnm = Trim(sGetVal)
                        .GetText 5, sRow, sGetVal:      cPis201.stkcd = Trim(sGetVal)
                        .GetText 7, sRow, sGetVal:      cPis201.lotno = Trim(sGetVal)
                        .GetText 8, sRow, sGetVal:      cPis201.expirydt = Format(sGetVal, "yyyyMMdd")
                        
                        .GetText 13, sRow, sGetVal:     cPis201.CD_COMPANY = Trim(sGetVal)
                        .GetText 14, sRow, sGetVal:     cPis201.CD_PLANT = Trim(sGetVal)
                        .GetText 15, sRow, sGetVal:     cPis201.CD_BIZAREA = Trim(sGetVal)
                        .GetText 16, sRow, sGetVal:     cPis201.NO_IO = Trim(sGetVal)
                        .GetText 17, sRow, sGetVal:     cPis201.NO_IOLINE = Val(sGetVal)
                        .GetText 27, sRow, sGetVal:     cPis201.NO_IOLINE2 = Val(sGetVal)
                        
                        .GetText 22, sRow, sGetVal:     cPis201.iqty_im = Val(sGetVal)
                        .GetText 23, sRow, sGetVal:     sRateQty = Val(sGetVal)
                        .GetText 24, sRow, sGetVal:     cPis201.unitamt = Val(sGetVal)
                        .GetText 25, sRow, sGetVal:     cPis201.amt = Val(sGetVal)
                        .GetText 26, sRow, sGetVal:     cPis201.taxamt = Val(sGetVal)
                        
                        If sRateQty = 0 Then
                            cPis201.iqty_so = cPis201.iqty_im
                        Else
                            cPis201.iqty_so = cPis201.iqty_im * sRateQty
                        End If
                        
                        Call cDb.csBegin
                        sReturn = cPis201.cfSave
                        If sReturn Then
                            Call cDb.csCommit
                            .SetText 1, sRow, ""
                            .SetText 12, sRow, "입고"
                        Else
                            Call cDb.csRollback
                            Exit For
                        End If
                    End If
                End If
            Next sRow
        End With
        
        If sReturn Then Call cmdFind_Click
    End If
    MousePointer = vbDefault

End Sub

Private Sub cmdFind_Click()
Dim sRow As Long, sFrDt As String, sToDt As String

    MousePointer = vbHourglass
    sFrDt = Format(dtpFrdt.Value, "yyyyMMdd")
    sToDt = Format(dtpTodt.Value, "yyyyMMdd")
    
    gSql = "SELECT A.*, X.NM_ITEM, X.UNIT_SO_FACT, B.LN_PARTNER AS NM_PARTNER FROM " & gTBLenter & " A                          " & vbNewLine & _
           "       LEFT JOIN " & gTBLstk & " X ON A.CD_COMPANY=X.CD_COMPANY AND A.CD_PLANT=X.CD_PLANT AND A.CD_ITEM=X.CD_ITEM   " & vbNewLine & _
           "       LEFT JOIN " & gTBLPartner & " B ON A.CD_COMPANY=B.CD_COMPANY AND A.CD_PARTNER=B.CD_PARTNER                   " & vbNewLine & _
           " WHERE A.CD_COMPANY='" & gCompany & "' AND A.CD_PLANT='" & gAreaCd & "'                                             " & vbNewLine & _
           "   AND A.DT_IO BETWEEN '" & sFrDt & "' AND '" & sToDt & "'                                                          " & vbNewLine
    If cboProc.ListIndex > 0 Then
        gSql = gSql & "  AND A.YN_LAB='" & IIf(cboProc.ListIndex = 1, "N", "Y") & "'                                            " & vbNewLine
    End If
    gSql = gSql & " ORDER BY A.DT_IO, A.NO_IOLINE"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                gPrgBar.Max = .RecordCount:     gPrgBar.Value = 0
                gPrgBar.Visible = True:         gPrgBar.Refresh
                Call gsSpreadClear(spList, .RecordCount, True)
                While (Not .EOF)
                    sRow = sRow + 1
                    
                    spList.SetText 1, sRow, ""
                    spList.SetText 2, sRow, Format("" & .Fields("DT_IO").Value, "####-##-##")
                    spList.SetText 3, sRow, "" & .Fields("CD_PARTNER").Value
                    spList.SetText 4, sRow, "" & .Fields("NM_PARTNER").Value
                    spList.SetText 5, sRow, "" & .Fields("CD_ITEM").Value
                    spList.SetText 6, sRow, "" & .Fields("NM_ITEM").Value
                    spList.SetText 7, sRow, "" & .Fields("NO_LOT").Value
                    spList.SetText 8, sRow, Format("" & .Fields("DT_LIMIT").Value, "####-##-##")
                    spList.SetText 9, sRow, gfQtyInputStr(Val("" & .Fields("QT_IO").Value))
                    spList.SetText 10, sRow, gfCurrencyStr(Val("" & .Fields("UM").Value))
                    spList.SetText 11, sRow, gfCurrencyStr(Val("" & .Fields("AM").Value))
                    spList.SetText 12, sRow, IIf("" & .Fields("YN_LAB").Value = "Y", "입고", "")
                    
                    spList.SetText 13, sRow, "" & .Fields("CD_COMPANY").Value
                    spList.SetText 14, sRow, "" & .Fields("CD_PLANT").Value
                    spList.SetText 15, sRow, "" & .Fields("CD_BIZAREA").Value
                    spList.SetText 16, sRow, "" & .Fields("NO_IO").Value
                    spList.SetText 17, sRow, "" & .Fields("NO_IOLINE").Value
                    spList.SetText 18, sRow, "" & .Fields("CD_SL").Value
                    spList.SetText 19, sRow, "" & .Fields("FG_IO").Value
                    spList.SetText 20, sRow, "" & .Fields("CD_QTIOTP").Value
                    spList.SetText 21, sRow, "" & .Fields("NO_EMP").Value
                    
                    spList.SetText 22, sRow, "" & .Fields("QT_IO").Value
                    spList.SetText 23, sRow, "" & .Fields("UNIT_SO_FACT").Value
                    spList.SetText 24, sRow, "" & .Fields("UM").Value
                    spList.SetText 25, sRow, "" & .Fields("AM").Value
                    spList.SetText 26, sRow, "" & .Fields("VAT").Value
                    
'                    spList.SetText 27, sRow, "" & .Fields("NO_IOLINE2").Value
                    
                    .MoveNext
                Wend
                gPrgBar.Visible = False
                
                Call gsButtonEnable(cmdFind, False)
                Call gsButtonEnable(cmdEnter, True)
                Call gsButtonEnable(cmdCancel, True)
            Else
                Call gsSpreadClear(spList, 0, True)
                MsgBox "해당기간에 입고된 자료가 없습니다.!", vbCritical
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

    grpMain.BorderColor = gGrpLineColor
    grpMain.BackColor = gGrpBackColor

    Me.Show
    
    With cboProc
        .AddItem "전체", 0
        .AddItem "미입고", 1
        .AddItem "입고", 2
        
        .ListIndex = 1
    End With
    
    dtpFrdt.Value = gfSystemDate
    dtpTodt.Value = gfSystemDate
    
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

