VERSION 5.00
Object = "{D8F5B61D-9152-4399-BF30-A1E4F3F072F6}#4.0#0"; "IGTabs40.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{94265C54-C72D-40E9-95C4-FBB6BF532C26}#2.0#0"; "XGroupBox.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form PIS103 
   BackColor       =   &H00FFFFFF&
   Caption         =   "공통코드관리"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9045
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
   ScaleHeight     =   8250
   ScaleWidth      =   9045
   WindowState     =   2  '최대화
   Begin XLibrary_XGroupBox.XGroupBox grpMain 
      Height          =   8175
      Left            =   30
      Top             =   30
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   14420
      BackColor       =   16777215
      BorderColor     =   10070188
      BorderRoundNum  =   0
      BorderStyle     =   1
      TextColor       =   0
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
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
      Begin BHButton.BHImageButton cmdSave 
         Height          =   375
         Left            =   5220
         TabIndex        =   6
         Top             =   7680
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
         TransparentPicture=   "PIS103.frx":0000
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdDelete 
         Height          =   375
         Left            =   6450
         TabIndex        =   8
         Top             =   7680
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
         TransparentPicture=   "PIS103.frx":17C2
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClose 
         Height          =   375
         Left            =   7680
         TabIndex        =   7
         Top             =   7680
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
         TransparentPicture=   "PIS103.frx":2F84
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin ActiveTabs.SSActiveTabs tabItem 
         Height          =   7470
         Left            =   90
         TabIndex        =   0
         Top             =   60
         Width           =   8805
         _ExtentX        =   15531
         _ExtentY        =   13176
         _Version        =   262144
         BackColor       =   16311512
         TabCount        =   2
         BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Tabs            =   "PIS103.frx":4746
         Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
            Height          =   7080
            Left            =   30
            TabIndex        =   2
            Top             =   360
            Width           =   8745
            _ExtentX        =   15425
            _ExtentY        =   12488
            _Version        =   262144
            TabGuid         =   "PIS103.frx":47C4
            Begin FPSpreadADO.fpSpread spReason 
               CausesValidation=   0   'False
               Height          =   6945
               Left            =   60
               TabIndex        =   4
               Tag             =   "20001"
               Top             =   60
               Width           =   8640
               _Version        =   524288
               _ExtentX        =   15240
               _ExtentY        =   12250
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
               MaxCols         =   4
               MaxRows         =   489
               Protect         =   0   'False
               ScrollBars      =   2
               SelectBlockOptions=   1
               ShadowColor     =   14737632
               ShadowDark      =   12632256
               SpreadDesigner  =   "PIS103.frx":47EC
               VisibleCols     =   3
               VisibleRows     =   10
               CellNoteIndicatorColor=   16576
            End
         End
         Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
            Height          =   7080
            Left            =   30
            TabIndex        =   1
            Top             =   360
            Width           =   8745
            _ExtentX        =   15425
            _ExtentY        =   12488
            _Version        =   262144
            TabGuid         =   "PIS103.frx":5263
            Begin FPSpreadADO.fpSpread spOper 
               CausesValidation=   0   'False
               Height          =   6945
               Left            =   60
               TabIndex        =   3
               Tag             =   "20001"
               Top             =   60
               Width           =   8610
               _Version        =   524288
               _ExtentX        =   15187
               _ExtentY        =   12250
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
               SpreadDesigner  =   "PIS103.frx":528B
               VisibleCols     =   3
               VisibleRows     =   10
               CellNoteIndicatorColor=   16576
            End
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "※ 기존 사용되는 코드는 삭제할 수 없습니다. ※"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   7830
         Width           =   4875
      End
   End
End
Attribute VB_Name = "PIS103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cPis005 As clsPis005, cPis006 As clsPis006

Private Sub psOperDisplay()
Dim sRow As Long

    With cPis005.cfList(True)
        If .State = adStateOpen Then
            spOper.Row = -1
            spOper.Col = 2:         spOper.TypeMaxEditLen = .Fields("OPERCD").DefinedSize
            spOper.Col = 3:         spOper.TypeMaxEditLen = .Fields("OPERNM").DefinedSize
            
            If Not .EOF Then
                Call gsSpreadClear(spOper, .RecordCount + 100, True)
                
                While (Not .EOF)
                    sRow = sRow + 1
                    
                    spOper.SetText 1, sRow, ""
                    spOper.SetText 2, sRow, "" & .Fields("OPERCD").Value
                    spOper.SetText 3, sRow, "" & .Fields("OPERNM").Value
                    
                    .MoveNext
                Wend
            Else
                Call gsSpreadClear(spOper, 100, True)
            End If
            .Close
        End If
    End With

End Sub

Private Sub psReasonDisplay()
Dim sRow As Long

    With cPis006.cfList(True)
        If .State = adStateOpen Then
            spReason.Row = -1
            spReason.Col = 3:         spReason.TypeMaxEditLen = .Fields("REASONCD").DefinedSize - 1
            spReason.Col = 4:         spReason.TypeMaxEditLen = .Fields("REASONNM").DefinedSize
        
            If Not .EOF Then
                Call gsSpreadClear(spReason, .RecordCount + 100, True)
                While (Not .EOF)
                    sRow = sRow + 1
                    
                    spReason.SetText 1, sRow, ""
                    spReason.SetText 3, sRow, Mid("" & .Fields("REASONCD").Value, 2)
                    spReason.SetText 4, sRow, "" & .Fields("REASONNM").Value
                    spReason.Row = sRow:      spReason.Col = 2
                    spReason.TypeComboBoxCurSel = Val("" & .Fields("KINDFG").Value)
                    
                    .MoveNext
                Wend
            Else
                Call gsSpreadClear(spReason, 100, True)
            End If
            .Close
        End If
    End With

End Sub

Private Sub psProcess(ByVal brSave As Boolean)
Dim sRow As Long, sReturn As Boolean, sGetVal As Variant

    sReturn = True
    If tabItem.Tabs(1).Selected Then
        With spOper
            For sRow = 1 To .MaxRows
                .GetText 2, sRow, sGetVal:      cPis005.opercd = sGetVal
                .GetText 1, sRow, sGetVal
                If Val(sGetVal) > 0 And Len(cPis005.opercd) > 0 Then
                    .GetText 3, sRow, sGetVal:  cPis005.opernm = sGetVal
                    
                    If brSave Then
                        sReturn = cPis005.cfSave
                    Else
                        sReturn = cPis005.cfDelete
                    End If
                    .SetText 1, sRow, ""
                End If
                
                If sReturn = False Then Exit For
            Next sRow
        End With
        
        If sReturn Then Call psOperDisplay
    Else
        With spReason
            For sRow = 1 To .MaxRows
                .GetText 3, sRow, sGetVal:      cPis006.reasoncd = sGetVal
                .GetText 1, sRow, sGetVal
                If Val(sGetVal) > 0 And Len(cPis006.reasoncd) > 0 Then
                    .GetText 4, sRow, sGetVal:  cPis006.reasonnm = sGetVal
                    .Row = sRow:        .Col = 2
                    If .TypeComboBoxCurSel < 0 Then
                        cPis006.kindfg = "0"
                    Else
                        cPis006.kindfg = .TypeComboBoxCurSel
                    End If
                    cPis006.reasoncd = cPis006.kindfg & cPis006.reasoncd
                    
                    If brSave Then
                        sReturn = cPis006.cfSave
                    Else
                        sReturn = cPis006.cfDelete
                    End If
                    .SetText 1, sRow, ""
                End If
                
                If sReturn = False Then Exit For
            Next sRow
        End With
        
        If sReturn Then Call psReasonDisplay
    End If

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdDelete_Click()

    MousePointer = vbHourglass
    If MsgBox("선택하신 자료를 삭제하시겠습니까 ?", vbYesNo + vbQuestion) = vbYes Then
        Call psProcess(False)
    End If
    MousePointer = vbDefault

End Sub

Private Sub cmdSave_Click()

    MousePointer = vbHourglass
    Call psProcess(True)
    MousePointer = vbDefault

End Sub

Private Sub Form_Activate()

    frmMain.lblTitle.Text = Me.Caption

End Sub

Private Sub Form_Load()

    Set cPis005 = New clsPis005
    Set cPis006 = New clsPis006

    grpMain.BorderColor = gGrpLineColor
    grpMain.BackColor = gGrpBackColor

    Me.Show
    
    spOper.Row = SpreadHeader:      spOper.Col = 1
    spOper.Text = "▒ 장비운영코드 현황 ▒"
    spOper.FontBold = True
    spOper.RowHeight(SpreadHeader) = 13
    spOper.RowHeight(SpreadHeader + 1) = 13

    spReason.Row = SpreadHeader:    spReason.Col = 1
    spReason.Text = "▒ 수기사유코드 현황 ▒"
    spReason.FontBold = True
    spReason.RowHeight(SpreadHeader) = 13
    spReason.RowHeight(SpreadHeader + 1) = 13
    
    With spReason
        .Row = -1
        .Col = 2
        .TypeComboBoxList = "공통사유" & vbTab & "장비운영사유" & vbTab & "수기검사사유" & vbTab & "수기출고사유" & vbTab & "유효기한변경"
    End With
    
    Call psOperDisplay
    Call psReasonDisplay
    
    tabItem.Tabs(1).Selected = True

End Sub

Private Sub Form_Resize()
On Error Resume Next

    grpMain.Top = (Me.ScaleHeight - grpMain.Height) / 2
    grpMain.Left = (Me.ScaleWidth - grpMain.Width) / 2

End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmMain.lblTitle.Text = ""
    
End Sub

Private Sub spOper_Change(ByVal Col As Long, ByVal Row As Long)

    With spOper
        If Row > 0 And Col > 1 Then
            .SetText 1, Row, 1
        End If
    End With

End Sub

Private Sub spReason_Change(ByVal Col As Long, ByVal Row As Long)

    With spReason
        If Row > 0 And Col > 1 Then
            .SetText 1, Row, 1
        End If
    End With

End Sub

