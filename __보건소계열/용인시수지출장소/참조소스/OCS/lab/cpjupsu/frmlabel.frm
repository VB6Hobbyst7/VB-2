VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmLabel 
   Caption         =   "Barcode Label Ãâ·ÂÈ­¸é"
   ClientHeight    =   3540
   ClientLeft      =   2505
   ClientTop       =   2475
   ClientWidth     =   6495
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "±¼¸²Ã¼"
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
   ScaleHeight     =   3540
   ScaleWidth      =   6495
   Begin Threed.SSPanel SSPanel2 
      Height          =   2805
      Left            =   4860
      TabIndex        =   5
      Top             =   180
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   4948
      _StockProps     =   15
      Caption         =   "SSPanel2"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Begin VB.TextBox txtPtno 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   90
         TabIndex        =   13
         Top             =   135
         Width           =   1230
      End
      Begin VB.TextBox txtSname 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   90
         TabIndex        =   12
         Top             =   420
         Width           =   1230
      End
      Begin VB.TextBox txtAgeYY 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   90
         TabIndex        =   11
         Top             =   705
         Width           =   1230
      End
      Begin VB.TextBox txtSex 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   90
         TabIndex        =   10
         Top             =   990
         Width           =   1230
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   90
         TabIndex        =   9
         Top             =   1275
         Width           =   1230
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   90
         TabIndex        =   8
         Top             =   1560
         Width           =   1230
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   90
         TabIndex        =   7
         Top             =   1845
         Width           =   1230
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   90
         TabIndex        =   6
         Top             =   2130
         Width           =   1230
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   2805
      Left            =   360
      TabIndex        =   2
      Top             =   180
      Width           =   4470
      _Version        =   65536
      _ExtentX        =   7885
      _ExtentY        =   4948
      _StockProps     =   15
      Caption         =   "SSPanel1"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Begin FPSpreadADO.fpSpread ssLabel 
         Height          =   2625
         Left            =   90
         TabIndex        =   3
         Top             =   90
         Width           =   4290
         _Version        =   196608
         _ExtentX        =   7567
         _ExtentY        =   4630
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         ScrollBars      =   2
         SpreadDesigner  =   "frmLabel.frx":0000
         Appearance      =   1
         ScrollBarTrack  =   1
      End
   End
   Begin VB.ListBox lstPrepare 
      BackColor       =   &H00FFC0C0&
      Height          =   2220
      Left            =   90
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   225
      Visible         =   0   'False
      Width           =   825
   End
   Begin MSForms.CommandButton cmdExit 
      Height          =   465
      Left            =   4860
      TabIndex        =   1
      Top             =   3015
      Width           =   1500
      Caption         =   "Á¾·á"
      PicturePosition =   327683
      Size            =   "2646;820"
      Picture         =   "frmLabel.frx":3C13
      FontName        =   "±¼¸²Ã¼"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdPrintOk 
      Height          =   465
      Left            =   3240
      TabIndex        =   0
      Top             =   3015
      Width           =   1590
      Caption         =   "Ãâ·ÂÈ®ÀÎ"
      PicturePosition =   327683
      Size            =   "2805;820"
      Picture         =   "frmLabel.frx":44ED
      FontName        =   "±¼¸²Ã¼"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()

End Sub

Private Sub cmdExit_Click()
    Unload Me
    
End Sub

Public Sub cmdPrintOk_Click()
    
    GoSub Label_Click_Ok
    
    If sLabelALLPrintIPD = "OK" Then
        DoEvents
        Unload frmLabel
    End If
    
    
    Exit Sub
    
Label_Click_Ok:
    Dim nCurrentRow     As Integer
    Dim sLabelText      As String
    
    For nCurrentRow = 1 To ssLabel.DataRowCnt
        ssLabel.Row = nCurrentRow
        ssLabel.Col = 4
        sLabelText = sLabelText & Trim(ssLabel.Text) & vbCrLf
    Next
    Debug.Print sLabelText
    'MsgBox sLabelText
    Return
    
End Sub

Private Sub Form_Activate()

    DoEvents
    GoSub Form_Locate_Set
    GoSub LstPrepare_Set
    GoSub SET_Label_Spread
    GoSub TextBox_Move
    
    If sLabelALLPrintIPD = "OK" Then
        DoEvents
        Call cmdPrintOk_Click
        'Unload Me
    End If
    
    
    
    Exit Sub
    
    
    
Form_Locate_Set:
    Me.Height = 4230
    Me.Width = 7125
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Return
    
    
LstPrepare_Set:
    Dim sBarText(100)    As String
    Dim nArraySeq        As Integer
    Dim sSLipno12        As String
    Dim nAryCnt          As Integer
    
    
    lstPrepare.Clear
    
    For i = 1 To frmMain.ssEnrol.DataRowCnt
        frmMain.ssEnrol.Row = i
        frmMain.ssEnrol.Col = 4
        If Len(Trim(frmMain.ssEnrol.Text)) > 2 Then
            frmMain.ssEnrol.Col = 2: sSLipno12 = frmMain.ssEnrol.Text
            frmMain.ssEnrol.Col = 6: sSLipno12 = sSLipno12 & Format(frmMain.ssEnrol.Text, "00000")
        End If
        
        nAryCnt = 0
        If False = isArrayText(sBarText, sSLipno12) Then
            sBarText(nAryCnt) = sSLipno12
            nAryCnt = nAryCnt + 1
        End If
    Next
    
    For j = 0 To 100
        If sBarText(j) = "" Then Exit For
        lstPrepare.AddItem sBarText(j)
    Next
    
    Return
    

SET_Label_Spread:
    Dim sSLipno     As String
    
    For i = 0 To lstPrepare.ListCount - 1
        sSLipno = Left(lstPrepare.List(i), 2)
        
        StrSql = ""
        StrSql = StrSql & " SELECT *"
        StrSql = StrSql & " FROM   TWEXAM_Specode"
        StrSql = StrSql & " WHERE  CodeGu  =  '12'"
        StrSql = StrSql & " AND    Codeky  =  '" & sSLipno & "'"
        
        If False = adoSetOpen(StrSql, adoSet) Then Return
        ssLabel.Row = ssLabel.DataRowCnt + 1
        ssLabel.Col = 1: ssLabel.Value = True
        ssLabel.Col = 2: ssLabel.Text = sSLipno
        ssLabel.Col = 3: ssLabel.Text = Mid(lstPrepare.List(i), 3, Len(lstPrepare.List(i)) - 2)
        ssLabel.Col = 4: ssLabel.Text = adoSet.Fields("Codenm").Value & ""
        ssLabel.Col = 5: ssLabel.TypeComboBoxCurSel = 0
        
        Call adoSetClose(adoSet)
        
    Next
    Return
    
TextBox_Move:
    Me.txtPtno.Text = frmMain.txtPtno.Text
    Me.txtSname.Text = frmMain.txtSname.Text
    Me.txtAgeYY.Text = frmMain.txtAgeYY.Text
    Me.txtSex.Text = frmMain.txtSex.Text
    
    Return

End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub
