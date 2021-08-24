VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{BD146989-F30B-4134-B202-680CC90638EF}#2.0#0"; "XTextBox.ocx"
Object = "{B88C4DC8-B707-435E-8B13-08058839823E}#2.0#0"; "XLabel.ocx"
Object = "{94265C54-C72D-40E9-95C4-FBB6BF532C26}#2.0#0"; "XGroupBox.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form PIS203 
   BackColor       =   &H00FFFFFF&
   Caption         =   "장비운영내역등록"
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
         TabIndex        =   10
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
         TransparentPicture=   "PIS203.frx":0000
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdSave 
         Height          =   375
         Left            =   11070
         TabIndex        =   9
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
         TransparentPicture=   "PIS203.frx":17C2
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdFind 
         Height          =   375
         Left            =   9840
         TabIndex        =   8
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
         TransparentPicture=   "PIS203.frx":2F84
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClear 
         Height          =   375
         Left            =   8610
         TabIndex        =   7
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
         TransparentPicture=   "PIS203.frx":4746
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdDelete 
         Height          =   375
         Left            =   12300
         TabIndex        =   6
         Top             =   840
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
         TransparentPicture=   "PIS203.frx":5F08
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin XLibrary_XTextBox.XTextBox lblMagam 
         Height          =   375
         Left            =   90
         TabIndex        =   5
         Top             =   840
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   661
         BackColor       =   16777215
         BorderColor     =   16744576
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "마감완료"
         BorderTextMargin=   4
         PasswordChar    =   ""
         MaxLength       =   0
         MouseCursor     =   4
         TextColor       =   255
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
         TextAlign       =   2
         Enabled         =   0   'False
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
         Begin TDBDate6Ctl.TDBDate dtpDt 
            Height          =   315
            Left            =   1590
            TabIndex        =   12
            Top             =   180
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   556
            Calendar        =   "PIS203.frx":76CA
            Caption         =   "PIS203.frx":77B1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "PIS203.frx":7814
            Keys            =   "PIS203.frx":7832
            Spin            =   "PIS203.frx":7890
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
         Begin BHButton.BHImageButton cmdEqp 
            Height          =   315
            Left            =   5550
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
            TransparentPicture=   "PIS203.frx":78B8
            BackColor       =   14737632
            AlphaColor      =   16777215
            ImgOutLineSize  =   3
         End
         Begin XLibrary_XTextBox.XTextBox txtEqpNm 
            Height          =   315
            Left            =   5910
            TabIndex        =   4
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
         Begin XLibrary_XLabel.XLabel XLabel1 
            Height          =   315
            Left            =   3060
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
         Begin XLibrary_XTextBox.XTextBox txtEqpcd 
            Height          =   315
            Left            =   4470
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
         Begin XLibrary_XLabel.XLabel XLabel4 
            Height          =   315
            Left            =   150
            TabIndex        =   1
            Top             =   180
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "운영일자"
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
         Height          =   8295
         Left            =   90
         TabIndex        =   0
         Tag             =   "20001"
         Top             =   1290
         Width           =   14640
         _Version        =   524288
         _ExtentX        =   25823
         _ExtentY        =   14631
         _StockProps     =   64
         BackColorStyle  =   1
         ButtonDrawMode  =   4
         ColHeaderDisplay=   0
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
         MaxCols         =   10
         MaxRows         =   489
         Protect         =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "PIS203.frx":907A
         VisibleCols     =   3
         VisibleRows     =   10
         CellNoteIndicatorColor=   16576
      End
   End
End
Attribute VB_Name = "PIS203"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cPis302 As clsPis302
Dim fOper() As String, fReason() As String

Private Sub psDataProcess(ByVal brSave As Boolean)
Dim sRow As Long, sGetVal As Variant, sReturn As Boolean

    With spList
        For sRow = 1 To .MaxRows
            .GetText 4, sRow, sGetVal:      cPis302.opercd = Trim(sGetVal)
            .GetText 1, sRow, sGetVal
            If Val(sGetVal) > 0 And Len(cPis302.opercd) > 0 Then
                cPis302.eqpcd = Trim(txtEqpcd.Text)
                cPis302.workdt = Format(dtpDt.Value, "yyyyMMdd")
                .GetText 2, sRow, sGetVal:      cPis302.seq = Val(sGetVal)
                .GetText 5, sRow, sGetVal:      cPis302.workcnt = Val(sGetVal)
                .GetText 7, sRow, sGetVal:      cPis302.reasoncd = Trim(sGetVal)
                .GetText 8, sRow, sGetVal:      cPis302.remark = Trim(sGetVal)
                cPis302.empid = gUserId
                cPis302.autofg = "0"
            
                If brSave Then
                    sReturn = cPis302.cfSave
                Else
                    sReturn = cPis302.cfDelete
                End If
                
                If sReturn Then
                    .SetText 1, sRow, ""
                Else
                    Exit For
                End If
            End If
        
        Next sRow
    End With
    
    If sReturn Then Call cmdFind_Click
    
End Sub

Private Sub cmdClear_Click()
Dim cPis005 As clsPis005, cPis006 As clsPis006, sStr As String, sRow As Integer

    sStr = "":      sRow = 0
    Set cPis005 = New clsPis005
    With cPis005.cfList(True)
        If .State = adStateOpen Then
            If Not .EOF Then
                ReDim fOper(.RecordCount) As String
                
                While (Not .EOF)
                    sStr = sStr & .Fields("OPERNM").Value
                    fOper(sRow) = "" & .Fields("OPERCD").Value
                    sRow = sRow + 1
                    
                    .MoveNext
                    If Not .EOF Then
                        sStr = sStr & vbTab
                    End If
                Wend
            End If
            .Close
        End If
    End With
    spList.Row = -1
    spList.Col = 3
    spList.TypeComboBoxList = sStr
    
    sStr = "":      sRow = 0
    Set cPis006 = New clsPis006
    With cPis006.cfList(True, gReasonMach)
        If .State = adStateOpen Then
            If Not .EOF Then
                ReDim fReason(.RecordCount) As String
                
                While (Not .EOF)
                    sStr = sStr & .Fields("REASONNM").Value
                    fReason(sRow) = "" & .Fields("REASONCD").Value
                    sRow = sRow + 1
                    
                    .MoveNext
                    If Not .EOF Then
                        sStr = sStr & vbTab
                    End If
                Wend
            End If
            .Close
        End If
    End With
    
    spList.Col = 6
    spList.TypeComboBoxList = sStr

    txtEqpcd.Text = ""
    txtEqpNm.Text = ""
    dtpDt.Value = gfSystemDate
    txtEqpcd.Enabled = True
    dtpDt.Enabled = True
    lblMagam.Text = ""

    Call gsSpreadClear(spList, 0, True)

    Call gsButtonEnable(cmdFind, True)
    Call gsButtonEnable(cmdSave, False)
    Call gsButtonEnable(cmdDelete, False)
    Call gsButtonEnable(cmdEqp, True)

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdDelete_Click()

    MousePointer = vbHourglass
    If MsgBox("선택하신 장비운영자료를 삭제하시겠습니까 ?", vbQuestion + vbYesNo) = vbYes Then
        Call psDataProcess(False)
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

Private Sub cmdFind_Click()
Dim sRow As Long
    
    If Len(txtEqpcd.Text) = 0 Then
        MsgBox "운영장비를 선택하세요.!", vbCritical
        txtEqpcd.SetFocus
        Exit Sub
    End If
    
    MousePointer = vbHourglass
    With cPis302.cfList(Format(dtpDt.Value, "yyyyMMdd"), Trim(txtEqpcd.Text))
        If .State = adStateOpen Then
            If Not .EOF Then
                Call gsSpreadClear(spList, .RecordCount + 50, True)
                While (Not .EOF)
                    sRow = sRow + 1
                    
                    spList.SetText 1, sRow, ""
                    spList.SetText 2, sRow, "" & .Fields("SEQ").Value
                    spList.SetText 3, sRow, "" & .Fields("OPERNM").Value
                    spList.SetText 4, sRow, "" & .Fields("OPERCD").Value
                    spList.SetText 5, sRow, "" & .Fields("WORKCNT").Value
                    spList.SetText 6, sRow, "" & .Fields("REASONNM").Value
                    spList.SetText 7, sRow, "" & .Fields("REASONCD").Value
                    spList.SetText 8, sRow, "" & .Fields("REMARK").Value
                    spList.SetText 9, sRow, "" & .Fields("EMPNM").Value
                    spList.SetText 10, sRow, "" & .Fields("MODDT").Value
                    
                    .MoveNext
                Wend
            Else
                Call gsSpreadClear(spList, 50, True)
            End If
            .Close
        End If
    End With
    
    If gfMagamCheck(Format(dtpDt.Value, "yyyyMMdd"), True) Then
        Call gsButtonEnable(cmdSave, True)
        Call gsButtonEnable(cmdDelete, True)
        lblMagam.Text = ""
    Else
        lblMagam.Text = "마감완료"
    End If
    
    Call gsButtonEnable(cmdFind, False)
    Call gsButtonEnable(cmdEqp, True)
    txtEqpcd.Enabled = False
    dtpDt.Enabled = False
    MousePointer = vbDefault
    
End Sub

Private Sub cmdSave_Click()

    MousePointer = vbHourglass
    Call psDataProcess(True)
    MousePointer = vbDefault

End Sub

Private Sub Form_Activate()

    frmMain.lblTitle.Text = Me.Caption

End Sub

Private Sub Form_Load()

    Set cPis302 = New clsPis302

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

Private Sub spList_Change(ByVal Col As Long, ByVal Row As Long)
Dim sGetVal As Variant

    If Row > 0 And Col > 1 Then
        spList.SetText 1, Row, 1

    End If

End Sub

Private Sub spList_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
Dim sCnt As Integer

    If Col = 6 Or Col = 3 Then
        spList.Row = Row
        spList.Col = Col
        sCnt = spList.TypeComboBoxCurSel
        
        spList.SetText Col + 1, Row, IIf(Col = 3, fOper(sCnt), fReason(sCnt))
    End If

End Sub

Private Sub txtEqpcd_LostFocus()

    txtEqpcd.Text = UCase(txtEqpcd.Text)
    txtEqpNm.Text = gfMachName(Trim(txtEqpcd.Text))
    If Len(txtEqpNm.Text) = 0 Then
        txtEqpcd.Text = ""
    End If

End Sub
