VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{94265C54-C72D-40E9-95C4-FBB6BF532C26}#2.0#0"; "XGroupBox.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form PIS902 
   BackColor       =   &H00FFFFFF&
   Caption         =   "검체저장고관리"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12405
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
   ScaleHeight     =   8490
   ScaleWidth      =   12405
   WindowState     =   2  '최대화
   Begin XLibrary_XGroupBox.XGroupBox grpMain 
      Height          =   8415
      Left            =   60
      Top             =   30
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   14843
      BackColor       =   16777215
      BorderColor     =   10070188
      BorderRoundNum  =   0
      BorderStyle     =   1
      TextColor       =   0
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
         Left            =   8580
         TabIndex        =   3
         Top             =   7950
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
         TransparentPicture=   "PIS902.frx":0000
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdDelete 
         Height          =   375
         Left            =   9810
         TabIndex        =   2
         Top             =   7950
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
         TransparentPicture=   "PIS902.frx":17C2
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClose 
         Height          =   375
         Left            =   11040
         TabIndex        =   1
         Top             =   7950
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
         TransparentPicture=   "PIS902.frx":2F84
         BackColor       =   12632319
         ImgOutLineSize  =   3
      End
      Begin FPSpreadADO.fpSpread spList 
         CausesValidation=   0   'False
         Height          =   7815
         Left            =   60
         TabIndex        =   0
         Tag             =   "20001"
         Top             =   60
         Width           =   12180
         _Version        =   524288
         _ExtentX        =   21484
         _ExtentY        =   13785
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
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "PIS902.frx":4746
         VisibleCols     =   3
         VisibleRows     =   10
         CellNoteIndicatorColor=   16576
      End
   End
End
Attribute VB_Name = "PIS902"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cPis092 As clsPis092

Private Sub psDisplay()
Dim sRow As Long

    With cPis092.cfList(True)
        If .State = adStateOpen Then
            spList.Row = -1
            spList.Col = 2:     spList.TypeMaxEditLen = .Fields("DEPOTCD").DefinedSize
            spList.Col = 3:     spList.TypeMaxEditLen = .Fields("DEPOTNM").DefinedSize
            spList.Col = 4:     spList.TypeMaxEditLen = .Fields("FLOOR").DefinedSize
            spList.Col = 5:     spList.TypeMaxEditLen = .Fields("RACKCNT").DefinedSize
            If Not .EOF Then
                Call gsSpreadClear(spList, .RecordCount + 100, True)
                While (Not .EOF)
                    sRow = sRow + 1
                    
                    spList.SetText 1, sRow, ""
                    spList.SetText 2, sRow, "" & .Fields("DEPOTCD").Value
                    spList.SetText 3, sRow, "" & .Fields("DEPOTNM").Value
                    spList.SetText 4, sRow, "" & .Fields("FLOOR").Value
                    spList.SetText 5, sRow, "" & .Fields("RACKCNT").Value
                    spList.SetText 6, sRow, "" & .Fields("USEFG").Value
                    spList.SetText 7, sRow, "" & .Fields("EMPNM").Value
                    spList.SetText 8, sRow, "" & .Fields("WRTDT").Value
                    spList.SetText 9, sRow, "" & .Fields("MODDT").Value
                    
                    .MoveNext
                Wend
            Else
                Call gsSpreadClear(spList, 100, True)
            End If
            .Close
        End If
    End With

End Sub

Private Sub psProcess(ByVal brJob As Boolean)
Dim sRow As Long, sGetVal As Variant, sReturn As Boolean

    sReturn = True
    With spList
        For sRow = 1 To .MaxRows
            .GetText 2, sRow, sGetVal:      cPis092.depotcd = sGetVal
            .GetText 1, sRow, sGetVal
            If Val(sGetVal) > 0 And Len(cPis092.depotcd) > 0 Then
                .GetText 3, sRow, sGetVal:  cPis092.depotnm = sGetVal
                .GetText 4, sRow, sGetVal:  cPis092.floor = Val(sGetVal)
                .GetText 5, sRow, sGetVal:  cPis092.rackcnt = Val(sGetVal)
                .GetText 6, sRow, sGetVal:  cPis092.usefg = Val(sGetVal)
                cPis092.empid = gUserId
                
                If brJob Then
                    sReturn = cPis092.cfSave
                Else
                    sReturn = cPis092.cfDelete
                End If
                
                If sReturn = False Then Exit For
            End If
        Next sRow
    End With
    
    If sReturn Then Call psDisplay

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

    Set cPis092 = New clsPis092
    
    grpMain.BorderColor = gGrpLineColor
    grpMain.BackColor = gGrpBackColor
    
    Me.Show
    
    Call psDisplay

End Sub

Private Sub Form_Resize()
On Error Resume Next

    grpMain.Top = (Me.ScaleHeight - grpMain.Height) / 2
    grpMain.Left = (Me.ScaleWidth - grpMain.Width) / 2
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmMain.lblTitle.Text = ""
    
End Sub

Private Sub spList_Change(ByVal Col As Long, ByVal Row As Long)

    If Row > 0 And Col > 1 Then
        With spList
            .SetText 1, Row, "1"
        End With
    End If

End Sub
