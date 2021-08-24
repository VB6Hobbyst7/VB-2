VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{BD146989-F30B-4134-B202-680CC90638EF}#2.0#0"; "XTextBox.ocx"
Object = "{B88C4DC8-B707-435E-8B13-08058839823E}#2.0#0"; "XLabel.ocx"
Object = "{94265C54-C72D-40E9-95C4-FBB6BF532C26}#2.0#0"; "XGroupBox.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form PIS207 
   BackColor       =   &H00FFFFFF&
   Caption         =   "창고불출등록(수기불출)"
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
         TabIndex        =   8
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
         TransparentPicture=   "PIS207.frx":0000
         BackColor       =   14737632
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdSave 
         Height          =   375
         Left            =   12300
         TabIndex        =   7
         Top             =   840
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "불출처리"
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
         TransparentPicture=   "PIS207.frx":17C2
         BackColor       =   14737632
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdFind 
         Height          =   375
         Left            =   11070
         TabIndex        =   6
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
         TransparentPicture=   "PIS207.frx":2F84
         BackColor       =   14737632
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClear 
         Height          =   375
         Left            =   9840
         TabIndex        =   5
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
         TransparentPicture=   "PIS207.frx":4746
         BackColor       =   16777215
         AlphaColor      =   16711680
         ImgOutLineSize  =   3
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
         Begin BHButton.BHImageButton cmdStk 
            Height          =   315
            Left            =   2700
            TabIndex        =   4
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TransparentPicture=   "PIS207.frx":5F08
            BackColor       =   14737632
            ImgOutLineSize  =   3
         End
         Begin XLibrary_XTextBox.XTextBox txtStkNm 
            Height          =   315
            Left            =   3060
            TabIndex        =   3
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
         Begin XLibrary_XLabel.XLabel XLabel4 
            Height          =   315
            Left            =   210
            TabIndex        =   2
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
            Left            =   1620
            TabIndex        =   1
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
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "PIS207.frx":649A
         VisibleCols     =   3
         VisibleRows     =   10
         CellNoteIndicatorColor=   16576
      End
   End
End
Attribute VB_Name = "PIS207"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()

    txtStkcd.Text = ""
    txtStkNm.Text = ""
    
    Call gsSpreadClear(spList, 0, True)
    
    Call gsButtonEnable(cmdFind, True)
    Call gsButtonEnable(cmdSave, False)
    
End Sub

Private Sub cmdClose_Click()

    Unload Me
        
End Sub

Private Sub cmdFind_Click()
Dim sRow As Long

    MousePointer = vbHourglass
    gSql = "SELECT A.*, X.NM_ITEM AS STKNM, C.CSTNM FROM S2PIS201 A                             " & vbNewLine & _
           "       LEFT JOIN " & gTBLstk & " X ON A.STKCD=X.CD_ITEM " & gERPStkCondition & "    " & vbNewLine & _
           "       LEFT JOIN S2PIS002 C ON A.CSTCD=C.CSTCD                                      " & vbNewLine & _
           " WHERE A.ENTDT <= '" & Format(gfSystemDate, "yyyyMMdd") & "'                        " & vbNewLine & _
           "   AND (A.OQTY_SO IS NULL OR A.OQTY_SO <= A.IQTY_SO)                                " & vbNewLine
    If Len(txtStkcd.Text) > 0 Then
        gSql = gSql & "  AND A.STKCD='" & Trim(txtStkcd.Text) & "'                              " & vbNewLine
    End If
    gSql = gSql & " ORDER BY A.ENTDT,A.STKCD, A.ENTSEQ"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                gPrgBar.Max = .RecordCount:     gPrgBar.Value = 0
                gPrgBar.Visible = True:         gPrgBar.Refresh
                Call gsSpreadClear(spList, .RecordCount, True)
                While (Not .EOF)
                    sRow = sRow + 1

                    spList.SetText 1, sRow, ""
                    spList.SetText 2, sRow, Format("" & .Fields("ENTDT").Value, "####-##-##")
                    spList.SetText 3, sRow, "" & .Fields("CSTCD").Value
                    spList.SetText 4, sRow, "" & .Fields("CSTNM").Value
                    spList.SetText 5, sRow, "" & .Fields("STKCD").Value
                    spList.SetText 6, sRow, "" & .Fields("STKNM").Value
                    spList.SetText 7, sRow, "" & .Fields("LOTNO").Value
                    spList.SetText 8, sRow, Format("" & .Fields("EXPIRYDT").Value, "####-##-##")
                    spList.SetText 9, sRow, gfQtyOutputStr(Val("" & .Fields("IQTY_SO").Value) - Val("" & .Fields("OQTY_SO").Value))
                    
'                    spList.SetText 10, sRow, gfQtyOutputStr(pfStkLotRemaind("" & .Fields("STKCD").Value, "" & .Fields("LOTNO").Value))
                    spList.SetText 11, sRow, ""
                    spList.SetText 12, sRow, ""
                    spList.SetText 13, sRow, "" & .Fields("ENTSEQ").Value
                    
                    .MoveNext
                Wend
                gPrgBar.Visible = False
                
                Call gsButtonEnable(cmdFind, False)
                Call gsButtonEnable(cmdSave, True)
            Else
                Call gsSpreadClear(spList, 0, True)
                MsgBox "자료가 없습니다.!", vbCritical
            End If
            .Close
        End If
    End With
    MousePointer = vbDefault
                    
End Sub

Private Function pfStkLotRemaind(ByVal brStk As String, ByVal brLot As String) As Long
Dim sReturn As Long

    gSql = "SELECT SUM(A.ENTQTY-A.USEQTY) AS RMDQTY FROM S2PIS401 A                 " & vbNewLine & _
           "       INNER JOIN S2PIS201 B ON A.ENTDT=B.ENTDT AND A.ENTSEQ=B.ENTSEQ   " & vbNewLine & _
           " WHERE A.STKCD='" & Trim(brStk) & "' AND B.LOTNO='" & Trim(brLot) & "'  " & vbNewLine & _
           " GROUP BY A.STKCD"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = Val("" & .Fields("RMDQTY").Value)
            End If
            .Close
        End If
    End With
    pfStkLotRemaind = sReturn

End Function

Private Sub cmdSave_Click()
Dim cPis401 As clsPis401
Dim sRow As Long, sReturn As Boolean, sGetVal As Variant

    MousePointer = vbHourglass
    Set cPis401 = New clsPis401
    If MsgBox("선택하신 자료를 불출처리하시겠습니까 ?", vbYesNo + vbQuestion) = vbYes Then
        With spList
            For sRow = 1 To .MaxRows
                .GetText 12, sRow, sGetVal:     cPis401.entqty = Val(sGetVal)
                .GetText 1, sRow, sGetVal
                If Val(sGetVal) > 0 And cPis401.entqty > 0 Then
                    .GetText 11, sRow, sGetVal
                    If Len(sGetVal) = 0 Then
                        cPis401.chuldt = Format(gfSystemDate, "yyyyMMdd")
                    Else
                        cPis401.chuldt = Format(sGetVal, "yyyyMMdd")
                    End If
                    cPis401.chulseq = 0
                    cPis401.empid = gUserId
                    .GetText 2, sRow, sGetVal:  cPis401.entdt = Format(sGetVal, "yyyyMMdd")
                    .GetText 5, sRow, sGetVal:  cPis401.stkcd = Trim(sGetVal)
                    .GetText 13, sRow, sGetVal: cPis401.entseq = Val(sGetVal)
                    
                    Call cDb.csBegin
                    sReturn = cPis401.cfSave
                    If sReturn Then
                        Call cDb.csCommit
                        .SetText 1, sRow, ""
                    Else
                        Call cDb.csRollback
                        Exit For
                    End If
                End If
            Next sRow
        End With
        
        If sReturn Then Call cmdFind_Click
    End If
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
Dim sGetVal As Variant, sRmdQty As Double

    If Row > 0 And Col > 1 Then
        spList.SetText 1, Row, 1
        If Col = 12 Then
            spList.GetText 9, Row, sGetVal:         sRmdQty = Val(Str(sGetVal))
            spList.GetText 12, Row, sGetVal
            If Val(sGetVal) > sRmdQty Then
                MsgBox "재고량보다 많은 수량을 불출할 수 없습니다.!", vbCritical
                spList.SetText 12, Row, sRmdQty
            End If
            
            spList.GetText 11, Row, sGetVal
            If Len(sGetVal) = 0 Then
                spList.SetText 11, Row, Format(gfSystemDate, "yyyy-MM-dd")
            End If
        End If
    End If
    
End Sub

Private Sub txtStkcd_LostFocus()

    txtStkNm.Text = gfStkName(Trim(txtStkcd.Text))
    If Len(txtStkNm.Text) = 0 Then txtStkcd.Text = ""

End Sub

