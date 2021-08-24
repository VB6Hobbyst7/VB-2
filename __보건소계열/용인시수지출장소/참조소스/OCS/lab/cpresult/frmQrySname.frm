VERSION 5.00
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmQrySname 
   Caption         =   "수진자명 조회"
   ClientHeight    =   4965
   ClientLeft      =   2835
   ClientTop       =   1725
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   7410
   Begin Threed.SSPanel SSPanel1 
      Height          =   555
      Left            =   90
      TabIndex        =   2
      Top             =   45
      Width           =   4650
      _Version        =   65536
      _ExtentX        =   8202
      _ExtentY        =   979
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Begin Threed.SSCommand cmdQrySname 
         Height          =   330
         Left            =   2610
         TabIndex        =   4
         Top             =   90
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   582
         _StockProps     =   78
         Caption         =   "조회확인"
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin VB.TextBox txtSname 
         Height          =   330
         Left            =   1080
         TabIndex        =   0
         Top             =   90
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "수진자명:"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   135
         Width           =   870
      End
   End
   Begin FPSpreadADO.fpSpread ssQrySname 
      Height          =   4155
      Left            =   90
      TabIndex        =   1
      Top             =   630
      Width           =   6990
      _Version        =   196608
      _ExtentX        =   12330
      _ExtentY        =   7329
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   7
      MaxRows         =   200
      ScrollBars      =   2
      SpreadDesigner  =   "frmQrySname.frx":0000
      Appearance      =   1
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmQrySname"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQrySname_Click()
    
    
    GoSub Search_Sname_Proc
    Exit Sub
'/--------------------------------------------------


Search_Sname_Proc:
'o  strSql = ""
'o  strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INDEX_PATIENT1) */"

    strSql = ""
    strSql = strSql & " SELECT a.*,"
    strSql = strSql & "        TO_Char(a.BirthDay, 'YYYY-MM-DD') BirthDay,"
    strSql = strSql & "        b.Jumin1, b.Jumin2, "
    strSql = strSql & "        c.DeptNamek, d.Drname"
    strSql = strSql & " FROM   TWEXAM_IDNOMST a,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PATIENT  b,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT     c,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR   d "
    strSql = strSql & " WHERE  a.Sname    LIKE  '" & txtSname.Text & "%'"
    strSql = strSql & " AND    a.Ptno     =     b.Ptno(+)"
    strSql = strSql & " AND    a.DeptCode =     c.DeptCode(+)"
    strSql = strSql & " AND    a.DrCode   =     d.Drcode(+)"
    
    Call SpreadSetClear(ssQrySname)
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    Do Until adoSet.EOF
        ssQrySname.Row = ssQrySname.DataRowCnt + 1
        ssQrySname.Col = 1: ssQrySname.Text = adoSet.Fields("Ptno").Value & ""
        ssQrySname.Col = 2: ssQrySname.Text = adoSet.Fields("Sname").Value & ""
        ssQrySname.Col = 3: ssQrySname.Text = adoSet.Fields("Sex").Value & ""
        ssQrySname.Col = 4: ssQrySname.Text = adoSet.Fields("Ageyy").Value & ""
        ssQrySname.Col = 5: ssQrySname.Text = adoSet.Fields("Jumin1").Value & "-" & _
                                              adoSet.Fields("Jumin2").Value & ""
        ssQrySname.Col = 6: ssQrySname.Text = adoSet.Fields("DeptNamek").Value & ""
        ssQrySname.Col = 7: ssQrySname.Text = adoSet.Fields("Drname").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return
    

End Sub

Private Sub Form_Load()
    
    txtSname.IMEMode = vbIMEModeHangul
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

Private Sub ssQrySname_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    If Row = 0 Then
        If Col > 0 Then
            GoSub Spread_Sort_Sub
        End If
    Else
        GoSub Data_Move_Set
        Unload Me
    End If
    Exit Sub
    


Spread_Sort_Sub:
    ssQrySname.Col = 1: ssQrySname.Col2 = ssQrySname.DataColCnt
    ssQrySname.Row = 1: ssQrySname.Row2 = ssQrySname.DataRowCnt
    
    ssQrySname.SortBy = SortByRow
    ssQrySname.SortKey(1) = Col
    
    If ssQrySname.SortKeyOrder(1) = SortKeyOrderDescending Then
        ssQrySname.SortKeyOrder(1) = SortKeyOrderAscending
    Else
        ssQrySname.SortKeyOrder(1) = SortKeyOrderDescending
    End If
    
    ssQrySname.Action = ActionSort
    
    Return
    

Data_Move_Set:
    ssQrySname.Row = Row
    ssQrySname.Col = 1
    Call SetWindowText(hWndReturn, ssQrySname.Text)
    
    Return
    
End Sub

Private Sub txtSname_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Trim(txtSname.Text) = "" Then Exit Sub
        Call cmdQrySname_Click
    End If
    
End Sub
