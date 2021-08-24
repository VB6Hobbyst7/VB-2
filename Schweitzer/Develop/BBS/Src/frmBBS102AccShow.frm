VERSION 5.00
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRCTL1.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS102AccShow 
   BorderStyle     =   3  
   Caption         =   "°ü·Ã°Ë»çÁ¶È¸"
   ClientHeight    =   7680
   ClientLeft      =   1080
   ClientTop       =   1425
   ClientWidth     =   10005
   Icon            =   "frmBBS102AccShow.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  
   MaxButton       =   0   
   MinButton       =   0   
   ScaleHeight     =   7680
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H80000005&
      Caption         =   "È®ÀÎ(&O)"
      Height          =   510
      Left            =   8640
      Style           =   1  
      TabIndex        =   4
      Top             =   7140
      Width           =   1320
   End
   Begin DRcontrol1.DrFrame fraLastRst 
      Height          =   1440
      Left            =   20
      TabIndex        =   0
      Top             =   0
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   2540
      Title           =   ""
      TitlePos        =   0
      DelLine         =   0
      BackColor       =   14411494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   
         Italic          =   0   
         Strikethrough   =   0   
      EndProperty
      Begin VB.TextBox Text1 
         Alignment       =   2  
         Appearance      =   0  
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1040
         MaxLength       =   10
         TabIndex        =   20
         Top             =   360
         Width           =   1400
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   300
         Index           =   0
         Left            =   40
         TabIndex        =   1
         Top             =   30
         Width           =   8340
         _ExtentX        =   14711
         _ExtentY        =   529
         BackColor       =   8388608
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   
            Italic          =   0   
            Strikethrough   =   0   
         EndProperty
         BorderStyle     =   0
         Caption         =   "¼öÇ÷¿¹Á¤È¯ÀÚ Á¤º¸"
         Appearance      =   0
         LeftGab         =   200
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   5
         Left            =   40
         TabIndex        =   11
         TabStop         =   0   
         Top             =   1050
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   
            Italic          =   0   
            Strikethrough   =   0   
         EndProperty
         Alignment       =   1
         Caption         =   "Ã³¹æ¸í"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   6
         Left            =   2480
         TabIndex        =   12
         TabStop         =   0   
         Top             =   705
         Width           =   1000
         _ExtentX        =   1773
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   
            Italic          =   0   
            Strikethrough   =   0   
         EndProperty
         Alignment       =   1
         Caption         =   "¿¬¶ôÃ³"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   4
         Left            =   2480
         TabIndex        =   13
         TabStop         =   0   
         Top             =   360
         Width           =   1000
         _ExtentX        =   1773
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   
            Italic          =   0   
            Strikethrough   =   0   
         EndProperty
         Alignment       =   1
         Caption         =   "ÁÖ¹Î¹øÈ£"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   2
         Left            =   40
         TabIndex        =   14
         TabStop         =   0   
         Top             =   705
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   
            Italic          =   0   
            Strikethrough   =   0   
         EndProperty
         Alignment       =   1
         Caption         =   "¼º¸í"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   1
         Left            =   40
         TabIndex        =   15
         TabStop         =   0   
         Top             =   360
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   
            Italic          =   0   
            Strikethrough   =   0   
         EndProperty
         Alignment       =   1
         Caption         =   "È¯ÀÚID"
         Appearance      =   0
      End
      Begin DRcontrol1.DrLabel lblDeptNm 
         Height          =   315
         Left            =   1040
         TabIndex        =   16
         TabStop         =   0   
         Top             =   1050
         Width           =   4450
         _ExtentX        =   7858
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   
            Italic          =   0   
            Strikethrough   =   0   
         EndProperty
         Caption         =   "Label1"
      End
      Begin DRcontrol1.DrLabel lblWard 
         Height          =   315
         Left            =   3500
         TabIndex        =   17
         TabStop         =   0   
         Top             =   705
         Width           =   2000
         _ExtentX        =   3519
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   
            Italic          =   0   
            Strikethrough   =   0   
         EndProperty
         Caption         =   "Label1"
      End
      Begin DRcontrol1.DrLabel lblSexAge 
         Height          =   315
         Left            =   3500
         TabIndex        =   18
         TabStop         =   0   
         Top             =   360
         Width           =   2000
         _ExtentX        =   3519
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   
            Italic          =   0   
            Strikethrough   =   0   
         EndProperty
         Caption         =   "Label1"
      End
      Begin DRcontrol1.DrLabel lblPtNm 
         Height          =   315
         Left            =   1040
         TabIndex        =   19
         TabStop         =   0   
         Top             =   705
         Width           =   1400
         _ExtentX        =   2461
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   
            Italic          =   0   
            Strikethrough   =   0   
         EndProperty
         Caption         =   "ÀÌ»ó´ë"
      End
      Begin VB.Label lblABO 
         Alignment       =   2  
         AutoSize        =   -1  
         BackStyle       =   0  
         Caption         =   "AB(AB)+"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   36
            Charset         =   129
            Weight          =   700
            Underline       =   0   
            Italic          =   0   
            Strikethrough   =   0   
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   720
         Left            =   5640
         TabIndex        =   22
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  
         Height          =   1030
         Left            =   5520
         TabIndex        =   21
         Top             =   345
         Width           =   2880
      End
   End
   Begin DRcontrol1.DrFrame DrFrame1 
      Height          =   1800
      Left            =   0
      TabIndex        =   2
      Top             =   1450
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   3175
      Title           =   ""
      TitlePos        =   0
      DelLine         =   0
      BackColor       =   14411494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   
         Italic          =   0   
         Strikethrough   =   0   
      EndProperty
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   300
         Index           =   2
         Left            =   40
         TabIndex        =   3
         Top             =   30
         Width           =   8340
         _ExtentX        =   14711
         _ExtentY        =   529
         BackColor       =   8388608
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   
            Italic          =   0   
            Strikethrough   =   0   
         EndProperty
         BorderStyle     =   0
         Caption         =   "¼öÇ÷³»¿ª"
         Appearance      =   0
         LeftGab         =   200
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   1335
         Left            =   40
         TabIndex        =   9
         Top             =   360
         Width           =   8340
         _Version        =   196608
         _ExtentX        =   14711
         _ExtentY        =   2355
         _StockProps     =   64
         AllowCellOverflow=   -1  
         AutoCalc        =   0   
         AutoClipboard   =   0   
         BackColorStyle  =   3
         DisplayColHeaders=   0   
         DisplayRowHeaders=   0   
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "µ¸¿òÃ¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   
            Italic          =   0   
            Strikethrough   =   0   
         EndProperty
         GridShowHoriz   =   0   
         GridShowVert    =   0   
         GridSolid       =   0   
         MaxCols         =   11
         OperationMode   =   1
         Protect         =   0   
         ScrollBars      =   2
         ShadowColor     =   12632256
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS102AccShow.frx":000C
         UnitType        =   0
         UserResize      =   0
         VisibleCols     =   8
         VisibleRows     =   22
         TextTip         =   4
      End
   End
   Begin DRcontrol1.DrFrame DrFrame2 
      Height          =   3855
      Left            =   0
      TabIndex        =   5
      Top             =   3270
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   6800
      Title           =   ""
      TitlePos        =   0
      DelLine         =   0
      BackColor       =   14411494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   
         Italic          =   0   
         Strikethrough   =   0   
      EndProperty
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   300
         Index           =   1
         Left            =   45
         TabIndex        =   6
         Top             =   30
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   529
         BackColor       =   8388608
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   
            Italic          =   0   
            Strikethrough   =   0   
         EndProperty
         BorderStyle     =   0
         Caption         =   "ÃÖ±Ù°ü·Ã°Ë»ç³»¿ª"
         Appearance      =   0
         LeftGab         =   200
      End
      Begin FPSpread.vaSpread tblResult 
         Height          =   3390
         Left            =   40
         TabIndex        =   7
         Top             =   360
         Width           =   4245
         _Version        =   196608
         _ExtentX        =   7488
         _ExtentY        =   5980
         _StockProps     =   64
         AllowCellOverflow=   -1  
         AutoCalc        =   0   
         AutoClipboard   =   0   
         BackColorStyle  =   3
         DisplayColHeaders=   0   
         DisplayRowHeaders=   0   
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "µ¸¿òÃ¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   
            Italic          =   0   
            Strikethrough   =   0   
         EndProperty
         GridShowHoriz   =   0   
         GridShowVert    =   0   
         GridSolid       =   0   
         MaxCols         =   11
         OperationMode   =   1
         Protect         =   0   
         ScrollBars      =   2
         ShadowColor     =   12632256
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS102AccShow.frx":1CC7
         UnitType        =   0
         UserResize      =   0
         VisibleCols     =   8
         VisibleRows     =   22
         TextTip         =   4
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   300
         Index           =   3
         Left            =   4200
         TabIndex        =   8
         Top             =   30
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   529
         BackColor       =   8388608
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   
            Italic          =   0   
            Strikethrough   =   0   
         EndProperty
         BorderStyle     =   0
         Caption         =   "Ab Screening ´©Àû°á°ú"
         Appearance      =   0
         LeftGab         =   200
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   3390
         Left            =   4320
         TabIndex        =   10
         Top             =   360
         Width           =   4125
         _Version        =   196608
         _ExtentX        =   7276
         _ExtentY        =   5980
         _StockProps     =   64
         AllowCellOverflow=   -1  
         AutoCalc        =   0   
         AutoClipboard   =   0   
         BackColorStyle  =   3
         DisplayColHeaders=   0   
         DisplayRowHeaders=   0   
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "µ¸¿òÃ¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   
            Italic          =   0   
            Strikethrough   =   0   
         EndProperty
         GridShowHoriz   =   0   
         GridShowVert    =   0   
         GridSolid       =   0   
         MaxCols         =   11
         OperationMode   =   1
         Protect         =   0   
         ScrollBars      =   2
         ShadowColor     =   12632256
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS102AccShow.frx":3982
         UnitType        =   0
         UserResize      =   0
         VisibleCols     =   8
         VisibleRows     =   22
         TextTip         =   4
      End
   End
End
Attribute VB_Name = "frmBBS102AccShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sTestDiv As String

Public Sub SpecialTest(ByVal sPtid As String, ByVal sPtNm As String, ByVal CboRel As ComboBox, ByVal TestDiv As String)
    Dim iTmx        As ListItem
    Dim strTmp      As String
    Dim ii          As Integer
    
    lvwLResult.ListItems.Clear
    
    sTestDiv = TestDiv
    
    For ii = 0 To CboRel.ListCount - 1
        
        With lvwLResult
            If medGetP(CboRel.List(ii), 3, vbTab) = sTestDiv Then
                strTmp = medGetP(CboRel.List(ii), 4, vbTab)
                Set iTmx = .ListItems.Add()
                iTmx.Text = medGetP(strTmp, 1, "-") & "-" & medGetP(strTmp, 2, "-") & "-" & medGetP(strTmp, 3, "-")
                iTmx.SubItems(1) = medGetP(CboRel.List(ii), 5, vbTab)
                iTmx.SubItems(2) = medGetP(strTmp, 4, "-")
               
            End If
        End With
    Next
    Call lvwLResult_ItemClick(iTmx)
    Set iTmx = Nothing
End Sub

Public Sub ComboDisplay(ByVal sTestcd As String, ByVal sCombo As String, ByRef objCombo As Object, _
                        ByRef objSpecial As Object, ByVal objMicro As Object)
    Dim ii As Integer
    Dim aryTmp() As String
    
    
    objSpecial.Visible = False
    objMicro.Visible = False
    
    If P_RealTestMicSpecial = False Then Exit Sub
    
    objCombo.Clear
    
    aryTmp = Split(sCombo, COL_DIV)
    
    For ii = LBound(aryTmp()) To UBound(aryTmp())
        If sTestcd = medGetP(aryTmp(ii), 2, vbTab) Then
            If medGetP(aryTmp(ii), 3, vbTab) = "1" Then objSpecial.Visible = True
            If medGetP(aryTmp(ii), 3, vbTab) = "2" Then objMicro.Visible = True
            
            objCombo.AddItem aryTmp(ii)
        End If
    Next
    If objCombo.ListCount = 0 Then
        objCombo.AddItem "< ¾øÀ½ >"
    End If
    objCombo.ListIndex = 0

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub



Private Sub Form_Load()
    lvwLResult.ListItems.Clear
    txtLastRst.TextRTF = ""
    medClearTable tblResult
    
End Sub



Private Sub lvwLResult_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Dim sWorkArea   As String
    Dim sAccDt      As String
    Dim sAccSeq     As String
    Dim sTestcd     As String
    Dim strTmp      As String
    
    DoEvents
    txtLastRst.TextRTF = ""
    medClearTable tblResult
    
    strTmp = Item.Text
    
    sWorkArea = medGetP(strTmp, 1, "-")
    sAccDt = medGetP(strTmp, 2, "-")
    sAccSeq = medGetP(strTmp, 3, "-")
    sTestcd = Item.SubItems(2)
    
    If sTestDiv = "1" Then
        Dim objETest    As New clsLISSpecialTest
        LisLabel7(2).Caption = "Æ¯¼ö°Ë»ç°á°ú(" & sWorkArea & "-" & sAccDt & "-" & sAccSeq & " º¸°íÀÏ½Ã : " & Item.SubItems(1) & " )"
        txtLastRst.TextRTF = objETest.GetResultText(sWorkArea, sAccDt, sAccSeq, sTestcd)
        Set objETest = New clsLISSpecialTest
    Else
        LisLabel7(1).Caption = "¹Ì»ý¹°°ü·Ã°Ë»ç°á°ú(" & sWorkArea & "-" & sAccDt & "-" & sAccSeq & " º¸°íÀÏ½Ã : " & Item.SubItems(1) & " )"
        Call DisplayMicroResult(sWorkArea, sAccDt, sAccSeq)
        
    End If

End Sub

Private Sub DisplayMicroResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer)
   
    
    Dim objResult   As New clsLISResultReview
    Dim i           As Integer
    Dim j           As Integer
    
    With objResult
      
        Call .ResultQuery(pWorkArea, pAccDt, pAccSeq)

        
        For i = 1 To .RstRow
            tblResult.Row = i   
            For j = 1 To 8
                tblResult.Col = j
                tblResult.ForeColor = .Get_ForeColor(j, i)
            Next
        Next
      
        
        tblResult.Row = 1
        tblResult.Row2 = tblResult.MaxRows
        tblResult.Col = 2
        tblResult.COL2 = tblResult.MaxCols
        tblResult.BlockMode = True
        tblResult.AllowCellOverflow = True
        tblResult.Clip = .ResultClipText 
        tblResult.BlockMode = False
      
        If .SortFg Then
            For i = 1 To .SensiCount
                tblResult.SortBy = SortByRow
                tblResult.SortKey(1) = 2  
                tblResult.SortKeyOrder(1) = SortKeyOrderAscending
                tblResult.Col = -1
                tblResult.Row = .AntiSortStartRow(i)   
                tblResult.Row2 = .AntiSortEndRow(i)    
                tblResult.Action = ActionSort
                tblResult.Row = .SortStartRow - 1 
                tblResult.Col = 2
                tblResult.FontUnderline = True
            Next
        Else
            tblResult.Col = 6
            tblResult.Row = -1
            tblResult.ForeColor = DCM_LightRed
            tblResult.FontBold = True
        End If
        If Val(.TestDiv) = TST_MicTest Then
            tblResult.Row = -1
            tblResult.Col = -1
            tblResult.BlockMode = True
            tblResult.AllowCellOverflow = True
            tblResult.TypeHAlign = TypeHAlignLeft
            tblResult.BlockMode = False
            tblResult.ColWidth(2) = 17
            For i = 1 To 5
                If .MicFg(i) Then
                    tblResult.ColWidth(i + 2) = 9
                Else
                    tblResult.ColWidth(i + 2) = 4
                End If
            Next
            tblResult.ColWidth(8) = 20
            tblResult.Col = 3: tblResult.COL2 = 7
            tblResult.Row = -1
            tblResult.BlockMode = True
            tblResult.FontBold = False
            tblResult.BlockMode = False
        Else
            tblResult.Row = 1: tblResult.Row2 = tblResult.MaxRows
            tblResult.Col = 3: tblResult.COL2 = 7
            tblResult.BlockMode = True
            tblResult.TypeHAlign = TypeHAlignCenter
            tblResult.BlockMode = False
            tblResult.ColWidth(2) = 13
            tblResult.ColWidth(3) = 9
            tblResult.ColWidth(4) = 9
            tblResult.ColWidth(5) = 3
            tblResult.ColWidth(6) = 5
            tblResult.ColWidth(7) = 13
        End If
        
        tblResult.Col = 1: tblResult.Row = 1
        
    End With
    
    Set objResult = Nothing
   
End Sub
