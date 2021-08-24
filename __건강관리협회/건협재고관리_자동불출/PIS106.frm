VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{BD146989-F30B-4134-B202-680CC90638EF}#2.0#0"; "XTextBox.ocx"
Object = "{B88C4DC8-B707-435E-8B13-08058839823E}#2.0#0"; "XLabel.ocx"
Object = "{94265C54-C72D-40E9-95C4-FBB6BF532C26}#2.0#0"; "XGroupBox.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form PIS106 
   BackColor       =   &H00FFFFFF&
   Caption         =   "장비운영일정"
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
         TabIndex        =   8
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
         TransparentPicture=   "PIS106.frx":0000
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
         TransparentPicture=   "PIS106.frx":17C2
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdFind 
         Height          =   375
         Left            =   9840
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
         TransparentPicture=   "PIS106.frx":2F84
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdSave 
         Height          =   375
         Left            =   11070
         TabIndex        =   5
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
         TransparentPicture=   "PIS106.frx":4746
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClose 
         Height          =   375
         Left            =   13530
         TabIndex        =   4
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
         TransparentPicture=   "PIS106.frx":5F08
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin FPSpreadADO.fpSpread spList 
         CausesValidation=   0   'False
         Height          =   8295
         Left            =   90
         TabIndex        =   3
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
         MaxCols         =   14
         MaxRows         =   489
         Protect         =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "PIS106.frx":76CA
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
         Begin BHButton.BHImageButton cmdEqp 
            Height          =   315
            Left            =   2670
            TabIndex        =   9
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
            TransparentPicture=   "PIS106.frx":8847
            BackColor       =   14737632
            AlphaColor      =   16777215
            ImgOutLineSize  =   3
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
         Begin XLibrary_XLabel.XLabel XLabel4 
            Height          =   315
            Left            =   180
            TabIndex        =   1
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
            Left            =   1590
            TabIndex        =   0
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
   End
End
Attribute VB_Name = "PIS106"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cPis103 As clsPis103, fOperCd() As String

Private Sub psDataProcess(ByVal brSave As Boolean)
Dim sRow As Long, sGetVal As Variant, sReturn As Boolean

    With spList
        cPis103.eqpcd = Trim(txtEqpcd.Text)
        cPis103.empid = gUserId
        
        For sRow = 1 To .MaxRows
            .GetText 4, sRow, sGetVal:      cPis103.opercd = Trim(sGetVal)
            .GetText 1, sRow, sGetVal
            If Val(sGetVal) > 0 And Len(cPis103.opercd) > 0 Then
                .GetText 2, sRow, sGetVal:      cPis103.seq = Val(sGetVal)
                .GetText 5, sRow, sGetVal:      cPis103.schnm = Trim(sGetVal)
                .Row = sRow
                .Col = 6:                       cPis103.cyclefg = .TypeComboBoxCurSel
                If cPis103.cyclefg = 1 Then
                    .Col = 7:                   cPis103.cycleday = .TypeComboBoxCurSel
                Else
                    .GetText 7, sRow, sGetVal:  cPis103.cycleday = Trim(sGetVal)
                End If
                .GetText 8, sRow, sGetVal:      cPis103.opercnt = Val(sGetVal)
                .GetText 9, sRow, sGetVal:      cPis103.startdt = Format(sGetVal, "yyyyMMdd")
                .GetText 10, sRow, sGetVal:     cPis103.enddt = Format(sGetVal, "yyyyMMdd")
                .GetText 11, sRow, sGetVal:     cPis103.pausefg = Trim(sGetVal)
                
                If brSave Then
                    sReturn = cPis103.cfSave
                Else
                    sReturn = cPis103.cfDelete
                End If
                
                If sReturn Then
                    .SetText sRow, 1, ""
                Else
                    Exit For
                End If
            End If
        Next sRow
    End With
    
    If sReturn Then Call cmdFind_Click
    
End Sub

Private Sub cmdClear_Click()

    txtEqpcd.Text = ""
    txtEqpNm.Text = ""
    
    Call gsSpreadClear(spList, 0, True)
    
    Call gsButtonEnable(cmdFind, True)
    Call gsButtonEnable(cmdEqp, True)
    Call gsButtonEnable(cmdSave, False)
    Call gsButtonEnable(cmdDelete, False)

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdDelete_Click()

    MousePointer = vbHourglass
    If MsgBox("선택하신 자료를 삭제하시겠습니까 ?", vbYesNo + vbQuestion) = vbYes Then
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

    MousePointer = vbHourglass
    If gWorkArea Then
        gSql = "SELECT A.*, B.OPERNM, C.EMPNM FROM S2PIS103 A                       " & vbNewLine & _
               "       LEFT JOIN S2PIS005 B ON A.OPERCD=B.OPERCD                    " & vbNewLine & _
               "       LEFT JOIN S2COM006 C ON A.EMPID=C.EMPID                      " & vbNewLine & _
               " WHERE A.EQPCD='" & Trim(txtEqpcd.Text) & "' ORDER BY A.EQPCD,A.SEQ"
    Else
        gSql = "SELECT A.*, B.OPERNM, C.USER_NM AS EMPNM FROM S2PIS103 A            " & vbNewLine & _
               "       LEFT JOIN S2PIS005 B ON A.OPERCD=B.OPERCD                    " & vbNewLine & _
               "       LEFT JOIN " & gKahpUserTable & " C ON A.EMPID=C.USERID       " & vbNewLine & _
               " WHERE A.EQPCD='" & Trim(txtEqpcd.Text) & "' ORDER BY A.EQPCD,A.SEQ"
    End If
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                Call gsSpreadClear(spList, .RecordCount + 50, True)
                While (Not .EOF)
                    sRow = sRow + 1
                    
                    spList.SetText 1, sRow, ""
                    spList.SetText 2, sRow, "" & .Fields("SEQ").Value
                    spList.SetText 3, sRow, "" & .Fields("OPERNM").Value
                    spList.SetText 4, sRow, "" & .Fields("OPERCD").Value
                    spList.SetText 5, sRow, "" & .Fields("SCHNM").Value
                    spList.Row = sRow
                    spList.Col = 6
                    spList.TypeComboBoxCurSel = Val("" & .Fields("CYCLEFG").Value)
                    Call spList_ComboSelChange(6, sRow)
                    If Val("" & .Fields("CYCLEFG").Value) = 1 Then
                        spList.Col = 7
                        spList.TypeComboBoxCurSel = Val("" & .Fields("CYCLEDAY").Value)
                    Else
                        spList.SetText 7, sRow, "" & .Fields("CYCLEDAY").Value
                    End If
                    spList.SetText 8, sRow, Val("" & .Fields("OPERCNT").Value)
                    spList.SetText 9, sRow, Format("" & .Fields("STARTDT").Value, "####-##-##")
                    spList.SetText 10, sRow, Format("" & .Fields("ENDDT").Value, "####-##-##")
                    spList.SetText 11, sRow, "" & .Fields("PAUSEFG").Value
                    spList.SetText 12, sRow, "" & .Fields("EMPNM").Value
                    spList.SetText 13, sRow, "" & .Fields("WRTDT").Value
                    spList.SetText 14, sRow, "" & .Fields("MODDT").Value
                    
                    .MoveNext
                Wend
            Else
                Call gsSpreadClear(spList, 50, True)
            End If
            .Close
        End If
    End With
    
    Call gsButtonEnable(cmdFind, False)
    Call gsButtonEnable(cmdEqp, False)
    Call gsButtonEnable(cmdSave, True)
    Call gsButtonEnable(cmdDelete, True)
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
Dim cPis005 As clsPis005, sStr As String, sRow As Integer

    Set cPis103 = New clsPis103
    
    grpMain.BorderColor = gGrpLineColor
    grpMain.BackColor = gGrpBackColor
    
    Me.Show
    
    Set cPis005 = New clsPis005
    With cPis005.cfList
        If .State = adStateOpen Then
            If Not .EOF Then
                ReDim fOperCd(.RecordCount) As String
                 
                While (Not .EOF)
                    sStr = sStr & .Fields("OPERNM").Value
                    fOperCd(sRow) = "" & .Fields("OPERCD").Value
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
    
    With spList
        .Row = -1
        .Col = 3:       .TypeComboBoxList = sStr
        
        sStr = "매일" & vbTab & "매주" & vbTab & "매월"
        .Col = 6:       .TypeComboBoxList = sStr
    End With
    
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

    If Row > 0 And Col > 1 Then
        spList.SetText 1, Row, 1
    End If

End Sub

Private Sub spList_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
Dim sStr As String, sSel As Integer

    With spList
        .Row = Row:     .Col = Col
        If Col = 3 Then
            sSel = .TypeComboBoxCurSel
            .SetText 4, Row, fOperCd(sSel)
        ElseIf Col = 6 Then
            Select Case .TypeComboBoxCurSel
                Case 0:
                        .Col = 7
                        .CellType = CellTypeStaticText
                        .Text = ""
                Case 1:
                        .Col = 7
                        .CellType = CellTypeComboBox
                        .TypeHAlign = TypeHAlignCenter
                        .TypeVAlign = TypeVAlignCenter
                        sStr = "일요일" & vbTab & "월요일" & vbTab & "화요일" & vbTab & "수요일" & vbTab & "목요일" & vbTab & "금요일" & vbTab & "토요일"
                        .TypeComboBoxList = sStr
                Case 2:
                        .Col = 7
                        .CellType = CellTypeNumber
                        .TypeHAlign = TypeHAlignCenter
                        .TypeVAlign = TypeVAlignCenter
                        .TypeNumberDecPlaces = 0
                        .TypeNumberMax = 31
                        .TypeNumberMin = 1
            End Select
        End If
    End With

End Sub

Private Sub txtEqpcd_LostFocus()

    txtEqpcd.Text = UCase(txtEqpcd.Text)
    txtEqpNm.Text = gfMachName(Trim(txtEqpcd.Text))
    If Len(txtEqpNm.Text) = 0 Then
        txtEqpcd.Text = ""
    End If

End Sub

