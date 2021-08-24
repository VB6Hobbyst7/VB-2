VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOdrDetail 
   Caption         =   "대기환자조회"
   ClientHeight    =   8100
   ClientLeft      =   2820
   ClientTop       =   660
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   5955
   Begin Threed.SSCommand cmdQryOk 
      Height          =   330
      Left            =   4005
      TabIndex        =   3
      Top             =   450
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "조회확인"
   End
   Begin MSComCtl2.DTPicker dtTdate 
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Top             =   450
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   25362435
      CurrentDate     =   36413
   End
   Begin MSComCtl2.DTPicker dtFdate 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   450
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   25362435
      CurrentDate     =   36413
   End
   Begin FPSpreadADO.fpSpread ssDetail 
      Height          =   6855
      Left            =   90
      TabIndex        =   0
      Top             =   1035
      Width           =   5685
      _Version        =   196608
      _ExtentX        =   10028
      _ExtentY        =   12091
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
      MaxCols         =   5
      MaxRows         =   200
      ScrollBars      =   2
      SpreadDesigner  =   "frmOdrDetail.frx":0000
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Caption         =   "From/To"
      Height          =   195
      Left            =   135
      TabIndex        =   5
      Top             =   495
      Width           =   870
   End
   Begin VB.Label Label1 
      Caption         =   "Order전송일자"
      Height          =   195
      Left            =   135
      TabIndex        =   4
      Top             =   225
      Width           =   2310
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmOdrDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQryOk_Click()
    Dim sFdate      As String
    Dim sTdate      As String
    
    sFdate = Format(dtFdate.Value, "yyyy-MM-dd")
    sTdate = Format(dtTdate.Value, "yyyy-MM-dd")
    
    Call Spread_Set_Clear(ssDetail)
    
    StrSql = ""
    StrSql = StrSql & " SELECT DISTINCT TO_CHAR(a.JeobsuDt,'YYYY-MM-DD') JeobsuDt , "
    StrSql = StrSql & "        a.Ptno, b.Sname, c.DeptNamek, a.GBER"
    StrSql = StrSql & " FROM   TWEXAM_ORDER  a,"
    StrSql = StrSql & "        TWBAS_Patient b,"
    StrSql = StrSql & "        TWBAS_Dept    c "
    StrSql = StrSql & " WHERE (a.JeobsuYn  = ' ' OR a.JeobsuYn IS NULL)"
    StrSql = StrSql & " AND    a.JeobsuDt >= TO_DATE('" & sFdate & "','YYYY-MM-DD')"
    StrSql = StrSql & " AND    a.JeobsuDt <= TO_DATE('" & sTdate & "','YYYY-MM-DD')"
    StrSql = StrSql & " AND    a.SLipno1  >  0"
    StrSql = StrSql & " AND    a.SLipno1  <  50"
    StrSql = StrSql & " AND    a.Ptno      = b.Ptno(+)"
    StrSql = StrSql & " AND    a.DeptCode  = c.DeptCode(+)"
    StrSql = StrSql & " ORder  by JeobsuDt Desc, a.Ptno"
    
    If False = adoSetOpen(StrSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        If ssDetail.Row = ssDetail.MaxRows Then
            ssDetail.MaxRows = ssDetail.MaxRows + 1
            ssDetail.RowHeight(ssDetail.MaxRows) = 10.5
        End If
        ssDetail.Row = ssDetail.DataRowCnt + 1
        ssDetail.Col = 1: ssDetail.Text = adoSet.Fields("JeobsuDt").Value & ""
        ssDetail.Col = 2: ssDetail.Text = adoSet.Fields("Ptno").Value & ""
        ssDetail.Col = 3: ssDetail.Text = adoSet.Fields("Sname").Value & ""
        ssDetail.Col = 4: ssDetail.Text = adoSet.Fields("DeptNamek").Value & ""
        ssDetail.Col = 5: ssDetail.Text = adoSet.Fields("GbEr").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
        

End Sub

Private Sub Form_Load()
    
    dtFdate.Value = Dual_Date_Get("yyyy-MM-dd")
    dtTdate.Value = Dual_Date_Get("yyyy-MM-dd")
    
    DoEvents
    Call cmdQryOk_Click
    
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

Private Sub ssDetail_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    If Row = 0 Then
        GoSub Sort_Spread_Sub
    Else
        ssDetail.Row = Row
        ssDetail.Col = 2
        frmMain.txtIDno.Text = ssDetail.Text
        DoEvents
        Unload Me
    End If
    
    Exit Sub
    
Sort_Spread_Sub:
    ssDetail.Col = 1
    ssDetail.Col2 = ssDetail.MaxCols
    ssDetail.Row = 1
    ssDetail.Row2 = ssDetail.DataRowCnt
    
    ssDetail.SortBy = SS_SORT_BY_ROW
    ssDetail.SortKey(1) = Col
    If ssDetail.SortKeyOrder(1) = SortKeyOrderDescending Then
        ssDetail.SortKeyOrder(1) = SortKeyOrderAscending
    Else
        ssDetail.SortKeyOrder(1) = SortKeyOrderDescending
    End If
    ssDetail.Action = SS_ACTION_SORT
    Return
    
    
End Sub

Private Sub ssDetail_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Call ssDetail_DblClick(1, ssDetail.ActiveRow)
    End If
    
End Sub
