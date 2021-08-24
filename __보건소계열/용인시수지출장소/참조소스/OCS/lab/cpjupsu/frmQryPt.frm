VERSION 5.00
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmQryPt 
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
   Begin VB.Frame Frame1 
      Caption         =   "일자별조건 확인Box"
      Height          =   825
      Left            =   225
      TabIndex        =   2
      Top             =   90
      Width           =   5280
      Begin MSComCtl2.DTPicker dtFdate 
         Height          =   330
         Left            =   225
         TabIndex        =   3
         Top             =   315
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24510467
         CurrentDate     =   36413
      End
      Begin MSComCtl2.DTPicker dtTdate 
         Height          =   330
         Left            =   1665
         TabIndex        =   4
         Top             =   315
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24510467
         CurrentDate     =   36413
      End
      Begin MSForms.CommandButton cmdQryOk 
         Height          =   555
         Left            =   3150
         TabIndex        =   0
         Top             =   180
         Width           =   1500
         Caption         =   "조회확인"
         PicturePosition =   327683
         Size            =   "2646;979"
         Picture         =   "frmQryPt.frx":0000
         FontName        =   "굴림"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
   End
   Begin FPSpreadADO.fpSpread ssDetail 
      Height          =   6855
      Left            =   90
      TabIndex        =   1
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
      SpreadDesigner  =   "frmQryPt.frx":08DA
      Appearance      =   1
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmQryPt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQryOK_Click()
    Dim sFdate      As String
    Dim sTdate      As String
    
    sFdate = Format(dtFdate.Value, "yyyy-MM-dd")
    sTdate = Format(dtTdate.Value, "yyyy-MM-dd")
    
    Call Spread_Set_Clear(ssDetail)
    
    strSql = ""
    strSql = strSql & " SELECT DISTINCT TO_CHAR(a.JeobsuDt,'YYYY-MM-DD') JeobsuDt ,     " & vbLf
    strSql = strSql & "        a.Ptno, b.Sname, c.DeptNamek, a.GBER                     " & vbLf
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Order  a,                             " & vbLf
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PATIENT b,                             " & vbLf
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT    c                              " & vbLf
    strSql = strSql & " WHERE (a.JeobsuYn  = ' ' OR a.JeobsuYn IS NULL)                 " & vbLf
    strSql = strSql & " AND    a.JeobsuDt >= TO_DATE('" & sFdate & "','YYYY-MM-DD')     " & vbLf
    strSql = strSql & " AND    a.JeobsuDt <= TO_DATE('" & sTdate & "','YYYY-MM-DD')     " & vbLf
    
    If GstrIOGubun = "OPD" Then
        strSql = strSql & " AND  a.GbIO    = 'O'                                        " & vbLf
    Else
        strSql = strSql & " AND  a.GbIO    = 'I'                                        " & vbLf
    End If
    strSql = strSql & " AND    a.SLipno1  >  0                                          " & vbLf
'C    strSql = strSql & " AND    a.SLipno1  <  52                                         " & vbLf          'Histology,Cytology 처리 안함
    strSql = strSql & " AND    a.SLipno1  <  90                                         " & vbLf          'Histology,Cytology 처리 안함
    strSql = strSql & " AND    a.Ptno      = b.Ptno(+)                                  " & vbLf
    strSql = strSql & " AND    a.DeptCode  = c.DeptCode(+)                              " & vbLf
    strSql = strSql & " ORder  by JeobsuDt Desc, a.Ptno                                 " & vbLf
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        If ssDetail.Row = ssDetail.MaxRows Then
            ssDetail.MaxRows = ssDetail.MaxRows + 1
            ssDetail.RowHeight(ssDetail.MaxRows) = 10.5
        End If
        ssDetail.Row = ssDetail.DataRowCnt + 1
        ssDetail.Col = 1: ssDetail.Text = adoSet.Fields("JeobsuDt").Value & "           " & vbLf
        ssDetail.Col = 2: ssDetail.Text = adoSet.Fields("Ptno").Value & "               " & vbLf
        ssDetail.Col = 3: ssDetail.Text = adoSet.Fields("Sname").Value & "              " & vbLf
        ssDetail.Col = 4: ssDetail.Text = adoSet.Fields("DeptNamek").Value & "          " & vbLf
        ssDetail.Col = 5: ssDetail.Text = adoSet.Fields("GbEr").Value & "               " & vbLf
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
        

End Sub

Private Sub Form_Load()
    
    dtFdate.Value = Dual_Date_Get("yyyy-MM-dd")
    dtTdate.Value = Dual_Date_Get("yyyy-MM-dd")
    
    DoEvents
    Call cmdQryOK_Click
    
    
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
        Call SetWindowText(hWndReturn, ssDetail.Text)
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
