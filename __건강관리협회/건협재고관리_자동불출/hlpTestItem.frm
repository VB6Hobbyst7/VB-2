VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{BD146989-F30B-4134-B202-680CC90638EF}#2.0#0"; "XTextBox.ocx"
Object = "{B88C4DC8-B707-435E-8B13-08058839823E}#2.0#0"; "XLabel.ocx"
Object = "{94265C54-C72D-40E9-95C4-FBB6BF532C26}#2.0#0"; "XGroupBox.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form hlpTestItem 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "검사코드선택"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5445
   ClipControls    =   0   'False
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
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin XLibrary_XGroupBox.XGroupBox XGroupBox1 
      Height          =   465
      Left            =   30
      Top             =   5130
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   820
      BackColor       =   16311512
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
      Begin XLibrary_XLabel.XLabel XLabel4 
         Height          =   315
         Left            =   90
         TabIndex        =   2
         Top             =   60
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         BackColor       =   16311512
         Text            =   "검색어"
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
      Begin XLibrary_XTextBox.XTextBox txtSearch 
         Height          =   315
         Left            =   750
         TabIndex        =   1
         Top             =   60
         Width           =   2895
         _ExtentX        =   5106
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
   Begin BHButton.BHImageButton cmdConfirm 
      Height          =   465
      Left            =   3810
      TabIndex        =   3
      Top             =   5130
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   820
      Caption         =   "확인"
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
      TransparentPicture=   "hlpTestItem.frx":0000
      BackColor       =   12632319
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton cmdClose 
      Height          =   465
      Left            =   4620
      TabIndex        =   4
      Top             =   5130
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   820
      Caption         =   "취소"
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
      TransparentPicture=   "hlpTestItem.frx":17C2
      BackColor       =   12632319
      ImgOutLineSize  =   3
   End
   Begin FPSpreadADO.fpSpread spList 
      CausesValidation=   0   'False
      Height          =   5055
      Left            =   30
      TabIndex        =   0
      Tag             =   "20001"
      Top             =   30
      Width           =   5370
      _Version        =   524288
      _ExtentX        =   9472
      _ExtentY        =   8916
      _StockProps     =   64
      BackColorStyle  =   1
      ColHeaderDisplay=   0
      DisplayRowHeaders=   0   'False
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
      SpreadDesigner  =   "hlpTestItem.frx":2F84
      VisibleCols     =   3
      VisibleRows     =   10
      CellNoteIndicatorColor=   16576
   End
End
Attribute VB_Name = "hlpTestItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub psTestDisplay()
Dim sRow As Long

    If gWorkArea Then
        gSql = "SELECT TESTCD,TESTNM FROM S2LAB001 ORDER BY TESTCD"
    Else
        gSql = "SELECT ITEMCODE AS TESTCD,ITEMHNM AS TESTNM FROM " & gKahpUser & "TWMED_ITEM ORDER BY ITEMCODE"
    End If
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                Call gsSpreadClear(spList, .RecordCount, True)
                While (Not .EOF)
                    sRow = sRow + 1
                    spList.SetText 1, sRow, ""
                    spList.SetText 2, sRow, "" & .Fields("TESTCD").Value
                    spList.SetText 3, sRow, "" & .Fields("TESTNM").Value
                    
                    .MoveNext
                Wend
            End If
        End If
    End With
    
End Sub

Private Sub cmdClose_Click()

    gHelpCode = ""
    Unload Me

End Sub

Private Sub cmdConfirm_Click()
Dim sRow As Long, sGetVal As Variant
    
    gHelpCode = ""
    With spList
        For sRow = 1 To .MaxRows
            .GetText 1, sRow, sGetVal
            If Val(sGetVal) > 0 Then
                If Len(gHelpCode) > 0 Then gHelpCode = gHelpCode & "|"
                .GetText 2, sRow, sGetVal
                gHelpCode = gHelpCode & Trim(sGetVal)
            End If
        Next sRow
    End With
    
    Unload Me
    
End Sub

Private Sub Form_Activate()
Dim sCol As Integer
    
    With spList
        .UserColAction = UserColActionSort
        For sCol = 2 To .MaxCols
            .ColUserSortIndicator(sCol) = ColUserSortIndicatorAscending
        Next sCol
        .Row = SpreadHeader:        .Col = 1
        If Me.Tag = "one" Then
            .Text = "▒ 검사코드현황 ▒"
            .FontBold = True
            
            .Col = 1:   .ColHidden = True
            .ColWidth(2) = 12
        Else
            .Text = "▒ 검사코드현황(다중선택) ▒"
            .FontBold = True
            
            .Col = 1:   .ColHidden = False
            .ColWidth(2) = 8
        End If
    End With
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        gHelpCode = ""
        Unload Me
    ElseIf KeyAscii = vbKeyReturn And Me.Tag = "one" Then
        Call spList_DblClick(spList.ActiveCol, spList.ActiveRow)
    End If

End Sub

Private Sub Form_Load()
    
    Call gsMousePoint(Me)
    Call psTestDisplay
    
    Me.KeyPreview = True

End Sub

Private Sub spList_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim sRow As Long, sGetVal As Variant
    
    If Me.Tag = "one" Then
        spList.GetText 2, Row, sGetVal
        gHelpCode = Trim(sGetVal)
        
        Unload Me
    End If
    
End Sub

Private Sub txtSearch_Change()
Dim sRow As Long, sLen As Integer, sCode As String, sName As String, sGetVal As Variant

    With spList
        sLen = HLen(txtSearch.Text)
        For sRow = 1 To .MaxRows
            .GetText 2, sRow, sGetVal:      sCode = Trim(sGetVal)
            .GetText 3, sRow, sGetVal:      sName = Trim(sGetVal)
            If (txtSearch.Text = HLeft(sCode, sLen)) Or (txtSearch.Text = HLeft(sName, sLen)) Then
                .Row = sRow:    .Col = 1
                .Action = ActionActiveCell
                Exit For
            End If
        Next sRow
    End With

End Sub
