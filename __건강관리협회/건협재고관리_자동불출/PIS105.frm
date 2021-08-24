VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{BD146989-F30B-4134-B202-680CC90638EF}#2.0#0"; "XTextBox.ocx"
Object = "{B88C4DC8-B707-435E-8B13-08058839823E}#2.0#0"; "XLabel.ocx"
Object = "{94265C54-C72D-40E9-95C4-FBB6BF532C26}#2.0#0"; "XGroupBox.ocx"
Object = "{1A3A9E7F-34C1-4F5C-BD80-63FA100EC4A0}#2.0#0"; "XCombobox.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form PIS105 
   BackColor       =   &H00FFFFFF&
   Caption         =   "장비운영별 소요시약"
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
      Begin BHButton.BHImageButton cmdClose 
         Height          =   375
         Left            =   13530
         TabIndex        =   12
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
         TransparentPicture=   "PIS105.frx":0000
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdSave 
         Height          =   375
         Left            =   12300
         TabIndex        =   11
         Top             =   840
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "저 장"
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
         TransparentPicture=   "PIS105.frx":17C2
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdFind 
         Height          =   375
         Left            =   11070
         TabIndex        =   10
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
         TransparentPicture=   "PIS105.frx":2F84
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClear 
         Height          =   375
         Left            =   9840
         TabIndex        =   9
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
         TransparentPicture=   "PIS105.frx":4746
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAdd 
         Height          =   855
         Left            =   8430
         TabIndex        =   8
         Top             =   3930
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1508
         Caption         =   "◀"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS105.frx":5F08
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdDel 
         Height          =   855
         Left            =   8430
         TabIndex        =   7
         Top             =   5640
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1508
         Caption         =   "▶"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS105.frx":76CA
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin XLibrary_XGroupBox.XGroupBox grpFind 
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
         Begin BHButton.BHImageButton cmdEqp 
            Height          =   315
            Left            =   2670
            TabIndex        =   13
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
            TransparentPicture=   "PIS105.frx":8E8C
            BackColor       =   14737632
            AlphaColor      =   16777215
            ImgOutLineSize  =   3
         End
         Begin XLibrary_XComboBox.XComboBox cboOper 
            Height          =   315
            Left            =   7710
            TabIndex        =   6
            Top             =   180
            Width           =   2625
            _ExtentX        =   4630
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
         Begin XLibrary_XLabel.XLabel XLabel11 
            Height          =   315
            Left            =   6300
            TabIndex        =   5
            Top             =   180
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "장비운영"
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
         Begin XLibrary_XTextBox.XTextBox txtEqpcd 
            Height          =   315
            Left            =   1590
            TabIndex        =   4
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
         Begin XLibrary_XLabel.XLabel XLabel4 
            Height          =   315
            Left            =   180
            TabIndex        =   3
            Top             =   180
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "검사장비"
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
         Begin XLibrary_XTextBox.XTextBox txtEqpNm 
            Height          =   315
            Left            =   3030
            TabIndex        =   2
            Top             =   180
            Width           =   3105
            _ExtentX        =   5477
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
      End
      Begin FPSpreadADO.fpSpread spStk 
         CausesValidation=   0   'False
         Height          =   8325
         Left            =   9210
         TabIndex        =   1
         Tag             =   "20001"
         Top             =   1290
         Width           =   5490
         _Version        =   524288
         _ExtentX        =   9684
         _ExtentY        =   14684
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
         MaxCols         =   3
         MaxRows         =   489
         Protect         =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "PIS105.frx":A64E
         VisibleCols     =   3
         VisibleRows     =   10
         CellNoteIndicatorColor=   16576
      End
      Begin FPSpreadADO.fpSpread spList 
         CausesValidation=   0   'False
         Height          =   8325
         Left            =   90
         TabIndex        =   0
         Tag             =   "20001"
         Top             =   1290
         Width           =   8190
         _Version        =   524288
         _ExtentX        =   14446
         _ExtentY        =   14684
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
         MaxCols         =   6
         MaxRows         =   489
         Protect         =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "PIS105.frx":B07B
         VisibleCols     =   3
         VisibleRows     =   10
         CellNoteIndicatorColor=   16576
      End
   End
End
Attribute VB_Name = "PIS105"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub psStkList()
Dim sRow As Long
    
    gSql = "SELECT X.CD_ITEM AS STKCD, X.NM_ITEM AS STKNM FROM " & gTBLstk & " X " & vbNewLine & _
           " WHERE NOT EXISTS(SELECT B.STKCD FROM S2PIS102 B WHERE B.EQPCD='" & Trim(txtEqpcd.Text) & "'" & vbNewLine & _
           "   AND B.OPERCD='" & cboOper.Tag & "' AND X.CD_ITEM=B.STKCD) " & gERPStkCondition & vbNewLine & _
           " ORDER BY X.CD_ITEM"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                Call gsSpreadClear(spStk, .RecordCount, True)
                While (Not .EOF)
                    sRow = sRow + 1
                    
                    spStk.SetText 1, sRow, ""
                    spStk.SetText 2, sRow, "" & .Fields("STKCD").Value
                    spStk.SetText 3, sRow, "" & .Fields("STKNM").Value
                    
                    .MoveNext
                Wend
            Else
                Call gsSpreadClear(spStk, 0, True)
            End If
            .Close
        End If
    End With

End Sub

Private Sub cboOper_Click()
Dim sOper() As String

    If cboOper.ListIndex >= 0 Then
        sOper = Split(cboOper.Text, gCboSplitStr)
        cboOper.Tag = Trim(sOper(1))
    Else
        cboOper.Tag = ""
    End If
    
End Sub

Private Sub cmdAdd_Click()
Dim sRow As Long, sRowA As Long, sGetVal As Variant, sCode As String

    With spStk
        For sRow = 1 To .MaxRows
            .GetText 2, sRow, sGetVal:      sCode = Trim(sGetVal)
            .GetText 1, sRow, sGetVal
            If Val(sGetVal) > 0 And Len(sCode) > 0 Then
                spList.MaxRows = spList.MaxRows + 1
                sRowA = spList.ActiveRow
                spList.InsertRows sRowA, 1
                
                spList.SetText 1, sRowA, 1
                spList.SetText 2, sRowA, sCode
                .GetText 3, sRow, sGetVal
                spList.SetText 3, sRowA, Trim(sGetVal)
                spList.SetText 4, sRowA, 1
                spList.SetText 5, sRowA, 0
                
                .DeleteRows sRow, 1:        .MaxRows = .MaxRows - 1
                sRow = sRow - 1
            End If
        Next sRow
    End With

End Sub

Private Sub cmdClear_Click()

    txtEqpcd.Text = ""
    txtEqpNm.Text = ""
    cboOper.ListIndex = -1
    
    Call gsSpreadClear(spList, 0, True)
    Call gsSpreadClear(spStk, 0, True)
    
    Call gsButtonEnable(cmdFind, True)
    Call gsButtonEnable(cmdEqp, True)
    
    Call gsButtonEnable(cmdAdd, False)
    Call gsButtonEnable(cmdDel, False)
    Call gsButtonEnable(cmdSave, False)
    
    grpFind.Enabled = True
    
End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdDel_Click()
Dim sRow As Long, sRowA As Long, sGetVal As Variant, sCode As String

    With spList
        For sRow = 1 To .MaxRows
            .GetText 2, sRow, sGetVal:      sCode = Trim(sGetVal)
            .GetText 1, sRow, sGetVal
            If Val(sGetVal) > 0 And Len(sCode) > 0 Then
                spStk.MaxRows = spStk.MaxRows + 1
                sRowA = 1
                spStk.InsertRows sRowA, 1
                
                spStk.SetText 2, sRowA, sCode
                .GetText 3, sRow, sGetVal
                spStk.SetText 3, sRowA, Trim(sGetVal)
                
                .DeleteRows sRow, 1:        .MaxRows = .MaxRows - 1
                sRow = sRow - 1
            End If
        Next sRow
    End With
    
End Sub

Private Sub cmdFind_Click()
Dim sRow As Long
    
    If Len(txtEqpcd.Text) = 0 Or Len(cboOper.Tag) = 0 Then
        MsgBox "검사장비와 장비운영내역을 선택하세요.!", vbCritical
        Exit Sub
    End If
    
    MousePointer = vbHourglass
    gSql = "SELECT A.*, X.NM_ITEM AS STKNM FROM S2PIS102 A LEFT JOIN " & gTBLstk & " X ON A.STKCD=X.CD_ITEM " & vbNewLine & gERPStkCondition & _
           " WHERE A.EQPCD='" & Trim(txtEqpcd.Text) & "' AND A.OPERCD='" & cboOper.Tag & "' ORDER BY A.EQPCD,A.OPERCD,A.STKCD"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                Call gsSpreadClear(spList, .RecordCount + 100, True)
                While (Not .EOF)
                    sRow = sRow + 1
                    
                    spList.SetText 1, sRow, ""
                    spList.SetText 2, sRow, "" & .Fields("STKCD").Value
                    spList.SetText 3, sRow, "" & .Fields("STKNM").Value
                    spList.SetText 4, sRow, "" & .Fields("QTY").Value
                    spList.SetText 5, sRow, "" & .Fields("LOSS").Value
                    spList.SetText 6, sRow, "" & .Fields("MODDT").Value
                    
                    .MoveNext
                Wend
            Else
                Call gsSpreadClear(spList, 100, True)
            End If
            .Close
        End If
    End With
    
    Call psStkList
    
    Call gsButtonEnable(cmdFind, False)
    Call gsButtonEnable(cmdEqp, False)
    
    Call gsButtonEnable(cmdAdd, True)
    Call gsButtonEnable(cmdDel, True)
    Call gsButtonEnable(cmdSave, True)
    
    grpFind.Enabled = False
    
    MousePointer = vbDefault

End Sub

Private Sub cmdSave_Click()
Dim cPis102 As clsPis102, sReturn As Boolean, sRow As Long, sGetVal As Variant
    
    MousePointer = vbHourglass
    If MsgBox("입력하신 자료를 저장하시겠습니까 ?", vbYesNo + vbQuestion) = vbYes Then
        Set cPis102 = New clsPis102
        
        Call cDb.csBegin
        
        cPis102.eqpcd = Trim(txtEqpcd.Text)
        cPis102.opercd = Trim(cboOper.Tag)
        sReturn = cPis102.cfDeleteAll
        If sReturn Then
            With spList
                For sRow = 1 To .MaxRows
                    .SetText 1, sRow, ""
                    .GetText 2, sRow, sGetVal:      cPis102.stkcd = Trim(sGetVal)
                    If Len(cPis102.stkcd) > 0 Then
                        .GetText 4, sRow, sGetVal:      cPis102.qty = Val(sGetVal)
                        .GetText 5, sRow, sGetVal:      cPis102.loss = Val(sGetVal)
                        
                        sReturn = cPis102.cfSave
                        If sReturn = False Then Exit For
                    End If
                Next sRow
            End With
        End If
        If sReturn Then
            Call cDb.csCommit
            
            Call cmdFind_Click
            MsgBox "장비운영별 소요시약자료가 저장되었습니다.!", vbInformation
        Else
            Call cDb.csRollback
        End If
    End If
    MousePointer = vbDefault

End Sub

Private Sub cmdEqp_Click()

    hlpTestMach.Show vbModal
    
    If Len(gHelpCode) > 0 Then
        txtEqpcd.Text = gHelpCode
        txtEqpNm.Text = gfMachName(gHelpCode)
    Else
        txtEqpcd.Text = ""
        txtEqpNm.Text = ""
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
    
    With spList
        .Row = SpreadHeader
        .Col = 2
        .Text = "▒ 소요시약현황 ▒"
        .FontBold = True
    
        .UserColAction = UserColActionSort
        For sCol = 2 To .MaxCols
            .ColUserSortIndicator(sCol) = ColUserSortIndicatorAscending
        Next sCol
        
        .Row = -1
        .Col = 4
        .TypeNumberDecPlaces = gDecimalQtyO
    End With
    With spStk
        .Row = SpreadHeader
        .Col = 2
        .Text = "▒ 시약품목현황 ▒"
        .FontBold = True
    
        .UserColAction = UserColActionSort
        For sCol = 2 To .MaxCols
            .ColUserSortIndicator(sCol) = ColUserSortIndicatorAscending
        Next sCol
    End With
    
    Call gsOperCombo(cboOper, False)
    
    Call cmdClear_Click

End Sub

Private Sub Form_Resize()
On Error Resume Next

    grpMain.Left = (Me.ScaleWidth - grpMain.Width) / 2
    grpMain.Height = Me.ScaleHeight - 50
    spList.Height = (grpMain.Height - spList.Top) - 50
    spStk.Height = (grpMain.Height - spStk.Top) - 50
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmMain.lblTitle.Text = ""
    
End Sub

Private Sub txtEqpcd_LostFocus()

    txtEqpcd.Text = UCase(txtEqpcd.Text)
    txtEqpNm.Text = gfMachName(Trim(txtEqpcd.Text))
    If Len(txtEqpNm.Text) = 0 Then
        txtEqpcd.Text = ""
    End If

End Sub


