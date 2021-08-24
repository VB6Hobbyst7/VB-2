VERSION 5.00
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmOills 
   BackColor       =   &H8000000A&
   Caption         =   "상병조회"
   ClientHeight    =   5685
   ClientLeft      =   5460
   ClientTop       =   2910
   ClientWidth     =   6060
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
   ScaleHeight     =   5685
   ScaleWidth      =   6060
   Begin FPSpreadADO.fpSpread ssOiLLs 
      Height          =   4695
      Left            =   1935
      TabIndex        =   5
      Top             =   675
      Width           =   4020
      _Version        =   196608
      _ExtentX        =   7091
      _ExtentY        =   8281
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   3
      MaxRows         =   100
      ScrollBars      =   2
      SpreadDesigner  =   "frmOills.frx":0000
      Appearance      =   1
      TextTip         =   1
      ScrollBarTrack  =   1
   End
   Begin VB.ListBox lstiLLgrp 
      BackColor       =   &H00404000&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   4560
      ItemData        =   "frmOills.frx":0E88
      Left            =   45
      List            =   "frmOills.frx":0E8F
      TabIndex        =   4
      Top             =   675
      Width           =   1770
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   555
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   5910
      _Version        =   65536
      _ExtentX        =   10425
      _ExtentY        =   979
      _StockProps     =   15
      Caption         =   "외래상병      "
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Alignment       =   4
      Begin VB.TextBox txtSname 
         Enabled         =   0   'False
         Height          =   330
         Left            =   2385
         TabIndex        =   3
         Top             =   135
         Width           =   1320
      End
      Begin VB.TextBox txtPtno 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1035
         TabIndex        =   2
         Top             =   135
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "병록번호"
         Height          =   240
         Left            =   135
         TabIndex        =   1
         Top             =   180
         Width           =   825
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmOills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim sLstDept        As String * 5
    Dim sLstBdate       As String * 10
    
    lstiLLgrp.Clear
    
    txtPtno.Text = gOiLLQryPtno
    GoSub Select_Sname_Sub
    
    If False = IsAdmission(txtPtno.Text) Then
        SSPanel1.Caption = "외래상병      "
        GoSub Select_OiLLs_Sub
        If lstiLLgrp.ListCount > 0 Then
            lstiLLgrp.ListIndex = 0
        End If
        Call lstiLLgrp_DblClick
    Else
        SSPanel1.Caption = "입원상병      "
        GoSub Select_iiLLs_Sub
    End If
    
    
    Me.ssOiLLs.Col = 2
    Me.ssOiLLs.ColHidden = True
    Me.ssOiLLs.Col = 3
    Me.ssOiLLs.ColHidden = False
    
    
    Exit Sub
    
    
Select_Sname_Sub:
'o  strSql = ""
'o  strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INDEX_PATIENT0) */"

    strSql = ""
    strSql = strSql & " SELECT * "
    strSql = strSql & " FROM   TW_MIS_PMPA.TWBAS_PATIENT"
    strSql = strSql & " WHERE  Ptno  =  '" & txtPtno.Text & "'"
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    
    txtSname.Text = adoSet.Fields("Sname").Value & ""
    Call adoSetClose(adoSet)
    Return


    
Select_iiLLs_Sub1:
'o  strSql = ""
'o  strSql = strSql & " SELECT /*+ INDEX(TWOCS_oILLS INX_oILLS1) */"

    strSql = ""
    strSql = strSql & " SELECT A.ILLCODE, B.ILLNAMEK, B.ILLNAMEE"
    strSql = strSql & " FROM   TW_MIS_PMPA.TWIPD_MASTER A,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_ILLS   B"
    strSql = strSql & " WHERE  A.PTNO = '" & txtPtno.Text & "'"
    strSql = strSql & " AND    RPAD(A.ILLCODE, 6) = B.ILLCODE(+)"
    If False = adoSetOpen(strSql, adoSet) Then Return
    Do Until adoSet.EOF
        ssOiLLs.Row = ssOiLLs.DataRowCnt + 1
        ssOiLLs.Col = 1: ssOiLLs.Text = adoSet.Fields("iLLCode").Value & ""
        ssOiLLs.Col = 2: ssOiLLs.Text = adoSet.Fields("iLLNameK").Value & ""
        ssOiLLs.Col = 3: ssOiLLs.Text = adoSet.Fields("iLLNameE").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    If lstiLLgrp.ListCount > 0 Then      '입원상병이 1 이상일때..
        lstiLLgrp.ListIndex = 0          '맨 첫번째것 자동 Click
        Call lstiLLgrp_DblClick          'SendKey lstillgrp_DblClick
    End If
    
    Return
    

Select_OiLLs_Sub:
    
    lstiLLgrp.Clear
    
    
    strSql = ""
    strSql = strSql & " SELECT   TO_CHAR(a.Bdate, 'YYYY-MM-DD') Bdate, "
    strSql = strSql & "          a.Ptno, a.DeptCode, Count(a.Ptno)"
    strSql = strSql & " FROM       TW_MIS_OCS.TWOCS_OILLS a"
    strSql = strSql & " WHERE    a.Ptno   =  '" & txtPtno.Text & "'"
    strSql = strSql & " GROUP BY a.Ptno, a.Bdate, a.DeptCode"
    strSql = strSql & " ORDER BY a.Bdate DESC, a.DeptCode ASC"
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sLstDept = adoSet.Fields("DeptCode").Value & ""
        sLstBdate = adoSet.Fields("Bdate").Value & ""
        lstiLLgrp.AddItem sLstDept & " " & sLstBdate
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return
    
    
Select_iiLLs_Sub:
    
    lstiLLgrp.Clear
    
    
    strSql = ""
    strSql = strSql & " SELECT   TO_CHAR(a.Bdate, 'YYYY-MM-DD') Bdate, "
    strSql = strSql & "          a.Ptno, a.DeptCode, Count(a.Ptno)"
    strSql = strSql & " FROM       TW_MIS_OCS.TWOCS_iILLS a"
    strSql = strSql & " WHERE    a.Ptno   =  '" & txtPtno.Text & "'"
    strSql = strSql & " GROUP BY a.Ptno, a.Bdate, a.DeptCode"
    strSql = strSql & " ORDER BY a.Bdate DESC, a.DeptCode ASC"
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sLstDept = adoSet.Fields("DeptCode").Value & ""
        sLstBdate = adoSet.Fields("Bdate").Value & ""
        lstiLLgrp.AddItem sLstDept & " " & sLstBdate
        adoSet.MoveNext
    Loop
    
    Call adoSetClose(adoSet)
    
    If lstiLLgrp.ListCount > 0 Then      '입원상병이 1 이상일때..
        lstiLLgrp.ListIndex = 0          '맨 첫번째것 자동 Click
        Call lstiLLgrp_DblClick          'SendKey lstillgrp_DblClick
    End If
    
    Return
    
    
End Sub

Public Sub lstiLLgrp_DblClick()
    Dim sViewDept       As String
    Dim sViewBdate      As String
    
    sViewDept = Left(lstiLLgrp.Text, 5)
    sViewBdate = Mid(lstiLLgrp.Text, 7, 10)
    
    '/    Hint   /
'o  strSql = ""
'o  strSql = strSql & " SELECT /*+ INDEX(TWOCS_oILLS INX_oILLS1) */"

    strSql = ""
    strSql = strSql & " SELECT  a.iLLCode, b.iLLNameE, b.iLLNameK"
    If SSPanel1.Caption = "입원상병      " Then
        strSql = strSql & " FROM   TW_MIS_COS.TWOCS_iILLS a,"
    Else
        strSql = strSql & " FROM   TW_MIS_OCS.TWOCS_OILLS a,"
    End If
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_ILLS  b "
    strSql = strSql & " WHERE  a.Ptno     = '" & txtPtno.Text & "'"
    strSql = strSql & " AND    a.DeptCode = '" & sViewDept & "'"
    strSql = strSql & " AND    a.Bdate    =      TO_DATE('" & sViewBdate & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.iLLCode  = B.iLLCode(+)"
    
    ssOiLLs.MaxRows = 0
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    ssOiLLs.MaxRows = adoSet.RecordCount
    ssOiLLs.RowHeight(-1) = 11
    
    Do Until adoSet.EOF
        ssOiLLs.Row = ssOiLLs.DataRowCnt + 1
        ssOiLLs.Col = 1: ssOiLLs.Text = adoSet.Fields("iLLCode").Value & ""
        ssOiLLs.Col = 2: ssOiLLs.Text = adoSet.Fields("iLLNamee").Value & ""
        ssOiLLs.Col = 3: ssOiLLs.Text = adoSet.Fields("iLLNamek").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

Private Sub ssOiLLs_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    If Row > 0 Then Exit Sub
    
    ssOiLLs.ReDraw = False
    Select Case Col
        Case 2:
            ssOiLLs.Col = 2: ssOiLLs.ColHidden = True
            ssOiLLs.Col = 3: ssOiLLs.ColHidden = False
            ssOiLLs.ColWidth(3) = 30.75
        Case 3
            ssOiLLs.Col = 2: ssOiLLs.ColHidden = False
            ssOiLLs.Col = 3: ssOiLLs.ColHidden = True
            ssOiLLs.ColWidth(2) = 30.75
    End Select
    ssOiLLs.ReDraw = True
    
End Sub

Private Sub ssOiLLs_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim sEtext       As String
    Dim sHtext       As String
    
    ssOiLLs.Row = Row
    ssOiLLs.Col = 2: sEtext = ssOiLLs.Text
    ssOiLLs.Col = 3: sHtext = ssOiLLs.Text
    
    
    TipText = "영) " & Trim(sEtext) & vbCrLf & _
              "한) " & Trim(sHtext)
    MultiLine = True
    If Trim(TipText) <> "" Then
        ShowTip = True: End If
    
    
End Sub

