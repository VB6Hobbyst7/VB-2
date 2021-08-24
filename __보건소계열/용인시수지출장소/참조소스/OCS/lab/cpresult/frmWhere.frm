VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{48932A52-981F-101B-A7FB-4A79242FD97B}#2.0#0"; "TAB32X20.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmWhere 
   Caption         =   "조건별조회"
   ClientHeight    =   6825
   ClientLeft      =   6915
   ClientTop       =   3600
   ClientWidth     =   5505
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
   ScaleHeight     =   6825
   ScaleWidth      =   5505
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4770
      Top             =   5535
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWhere.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWhere.frx":031C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '위 맞춤
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   635
      ButtonWidth     =   1535
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Move"
            Key             =   "Left"
            Object.ToolTipText     =   "Move Data to Left"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin FPSpreadADO.fpSpread sprQryList 
      Height          =   5460
      Left            =   0
      TabIndex        =   5
      Top             =   1305
      Width           =   5460
      _Version        =   196608
      _ExtentX        =   9631
      _ExtentY        =   9631
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
      MaxCols         =   6
      MaxRows         =   200
      ScrollBars      =   2
      SpreadDesigner  =   "frmWhere.frx":0BF8
      UserResize      =   0
      Appearance      =   1
   End
   Begin Threed.SSCommand cmdQuery 
      Height          =   870
      Left            =   4095
      TabIndex        =   4
      Top             =   405
      Width           =   825
      _Version        =   65536
      _ExtentX        =   1455
      _ExtentY        =   1535
      _StockProps     =   78
      Caption         =   "조회확인"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TabproLib.vaTabPro tabWhere 
      Height          =   870
      Left            =   0
      OleObjectBlob   =   "frmWhere.frx":171C
      TabIndex        =   2
      Top             =   405
      Width           =   4065
      Begin VB.ComboBox cmbWard 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -17714
         Style           =   2  '드롭다운 목록
         TabIndex        =   3
         Top             =   -15659
         Width           =   1455
      End
      Begin VB.ComboBox cmbDept 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -18029
         Style           =   2  '드롭다운 목록
         TabIndex        =   6
         Top             =   -15704
         Width           =   1905
      End
      Begin VB.TextBox txtSname 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -17534
         TabIndex        =   1
         Top             =   -15734
         Width           =   1365
      End
      Begin VB.TextBox txtPtno 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -17249
         TabIndex        =   0
         Top             =   -15689
         Width           =   1230
      End
   End
End
Attribute VB_Name = "frmWhere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQuery_Click()
    Dim sJeobsuDt       As String
    Dim iSLipno1        As Integer
    
    sJeobsuDt = Format(frmResult.dtJeobsu.Value, "YYYY-MM-DD")
    iSLipno1 = Val(Left(frmResult.cmbSLip.Text, 2))
    Call SpreadSetClear(sprQryList)
    
    Select Case tabWhere.ActiveTab
        Case 0: GoSub SELECT_Query_Ptno
        Case 1: GoSub SELECT_Query_Sname
        Case 2: GoSub SELECT_Query_Dept
        Case 3: GoSub SELECT_Query_Ward
    End Select
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        sprQryList.Row = sprQryList.DataRowCnt + 1
        sprQryList.Col = 1: sprQryList.Text = Format(adoSet.Fields("SLipno2").Value & "", "00000")
        sprQryList.Col = 2: sprQryList.Text = adoSet.Fields("Sname").Value & ""
        sprQryList.Col = 3: sprQryList.Text = adoSet.Fields("DeptCode").Value & ""
        If adoSet.Fields("GBIO").Value & "" = "I" Then
            sprQryList.Col = 4: sprQryList.Text = adoSet.Fields("RoomCode").Value & ""
        Else
            If Trim(adoSet.Fields("Deptcode").Value & "") = "ME" Then
                sprQryList.Col = 4: sprQryList.Text = "종합건진"
            Else
                sprQryList.Col = 4: sprQryList.Text = "외래"
            End If
        End If
        Select Case adoSet.Fields("Status").Value & ""
            Case "C": sprQryList.Col = 5: sprQryList.Text = "결과완료"
            Case "P": sprQryList.Col = 5: sprQryList.Text = "부분결과"
            Case "X": sprQryList.Col = 5: sprQryList.Text = "이상Data"
            Case "R": sprQryList.Col = 5: sprQryList.Text = "접수중"
        End Select
        
        sprQryList.Col = 6: sprQryList.Text = adoSet.Fields("JDate").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Exit Sub
    
    
    

SELECT_Query_Ptno:
    strSql = ""
    strSql = strSql & " SELECT a.*, b.Sname, TO_CHAR(a.JeobsuDt, 'yyyy-MM-dd') JDate"
    strSql = strSql & " FROM   TWEXAM_General  a,"
    strSql = strSql & "        TWEXAM_Idnomst  b "
    strSql = strSql & " WHERE  a.Ptno     = '" & txtPtno.Text & "'"
    'strSql = strSql & " AND    a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.SLipno1  = " & iSLipno1
    strSql = strSql & " AND    a.PTno     = b.Ptno(+)"
    strSql = strSql & " Order By a.JeobsuDt DESC , a.SLipno2 DESC"
    Return
    
    
    
SELECT_Query_Sname:
    strSql = ""
    strSql = strSql & " SELECT a.*, b.Sname, TO_CHAR(a.JeobsuDt, 'yyyy-MM-dd') JDate"
    strSql = strSql & " FROM   TWEXAM_General  a,"
    strSql = strSql & "        TWEXAM_Idnomst  b "
    strSql = strSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.SLipno1  = " & iSLipno1
    strSql = strSql & " AND    a.PTno     = b.Ptno(+)"
    strSql = strSql & " AND    b.Sname    Like '" & txtSname.Text & "%'"
    strSql = strSql & " Order By a.SLipno2"
    Return
    
    
SELECT_Query_Dept:
    strSql = ""
    strSql = strSql & " SELECT a.*, b.Sname, TO_CHAR(a.JeobsuDt, 'yyyy-MM-dd') JDate"
    strSql = strSql & " FROM   TWEXAM_General  a,"
    strSql = strSql & "        TWEXAM_Idnomst  b "
    strSql = strSql & " WHERE  a.DeptCode = '" & Left(cmbDept.Text, 4) & "'"
    strSql = strSql & " AND    a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.SLipno1  = " & iSLipno1
    strSql = strSql & " AND    a.PTno     = b.Ptno(+)"
    strSql = strSql & " Order By a.SLipno2"

    Return
    
    
SELECT_Query_Ward:
'o  strSql = ""
'o  strSql = strSql & "   SELECT /*+ INDEX (TWBas_Room INDEX_Room0) */"

    strSql = ""
    strSql = strSql & " SELECT   c.WardCode, a.*, d.Sname"
    strSql = strSql & "   FROM   TWEXAM_GENERAL a,"
    strSql = strSql & "          TW_MIS_PMPA.TWBas_Room     b,"
    strSql = strSql & "          TW_MIS_PMPA.TWBas_Ward     c,"
    strSql = strSql & "          TWEXAM_IDnomst d"
    strSql = strSql & "   WHERE  a.JEOBSUDT = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
    strSql = strSql & "   AND    a.SLIPNO1  = " & iSLipno1
    strSql = strSql & "   AND    a.RoomCode = b.RoomCode(+)"
    strSql = strSql & "   AND    b.WardCode = c.WardCode(+)"
    strSql = strSql & "   AND    a.Ptno     = d.Ptno(+)"
    strSql = strSql & "   Order  by a.SLipno2"
    
    
    Return
    

End Sub

Private Sub Form_Load()
    Dim sJeobsuDt       As String
    Dim iSLipno1        As Integer
    Dim sDeptC          As String * 4
    Dim sWardC          As String * 6
    
    sJeobsuDt = Format(frmResult.dtJeobsu.Value, "YYYY-MM-DD")
    iSLipno1 = Val(Left(frmResult.cmbSLip.Text, 2))
    
    GoSub SELECT_Query_Dept
    GoSub SELECT_Query_Ward
    Exit Sub
    
    
    
SELECT_Query_Dept:
'o  strSql = ""
'o  strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_DEPT INDEX_DEPT0) */"

    strSql = ""
    strSql = strSql & " SELECT a.DeptCode, b.DeptNamek"
    strSql = strSql & " FROM   TWEXAM_General a,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT     b"
    strSql = strSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.Deptcode = b.Deptcode(+)"
    strSql = strSql & " GROUP  BY a.DeptCode, b.DeptNamek"
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        sDeptC = adoSet.Fields("DeptCode").Value & ""
        cmbDept.AddItem sDeptC & "." & Trim(adoSet.Fields("DeptNamek").Value & "")
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return
    
SELECT_Query_Ward:
'o  strSql = ""
'o  strSql = strSql & "  SELECT /*+ INDEX (TWBas_Room INDEX_Room0) */"
    
    strSql = ""
    strSql = strSql & " SELECT  c.WardCode, c.WardName"
    strSql = strSql & "  FROM   TWEXAM_General a,             "
    strSql = strSql & "         TW_MIS_PMPA.TWBas_Room     b, "
    strSql = strSql & "         TW_MIS_PMPA.TWBas_Ward     c  "
    strSql = strSql & "  WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
    strSql = strSql & "  AND    a.RoomCode = b.Roomcode(+) "
    strSql = strSql & "  AND    b.WardCode = c.WardCode(+) "
    strSql = strSql & "  GROUP  BY c.WardCode, c.WardName"
    
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    Do Until adoSet.EOF
        sWardC = adoSet.Fields("WardCode").Value & ""
        cmbWard.AddItem sWardC & "." & Trim(adoSet.Fields("WardName").Value & "")
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
        
    Return
    
    
End Sub

Private Sub sprQryList_DblClick(ByVal Col As Long, ByVal Row As Long)
        
    If Row = 0 Then
        GoSub Spread_Sort_Sub
        Exit Sub
    End If
    
    sprQryList.Row = Row
    sprQryList.Col = 1: frmResult.txtSLipno2.Text = sprQryList.Text
    sprQryList.Col = 6: frmResult.dtJeobsu.Value = sprQryList.Text
    
    Call frmResult.txtSLipno2_KeyDown(vbKeyReturn, 1)
    Unload Me
    Exit Sub
    


Spread_Sort_Sub:
    sprQryList.Col = 1
    sprQryList.Col2 = sprQryList.MaxCols
    sprQryList.Row = 1
    sprQryList.Row2 = sprQryList.DataRowCnt
    
    sprQryList.SortBy = SS_SORT_BY_ROW
    sprQryList.SortKey(1) = Col
    
    If sprQryList.SortKeyOrder(1) = SortKeyOrderAscending Then
        sprQryList.SortKeyOrder(1) = SortKeyOrderDescending
    Else
        sprQryList.SortKeyOrder(1) = SortKeyOrderAscending
    End If
    
    sprQryList.Action = SS_ACTION_SORT
    Return
    

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1: Unload Me
        Case 3: GoSub LEFT_Move_Sub
    End Select
    Exit Sub
    
LEFT_Move_Sub:
    Dim sLabno      As String
    
    frmResult.lstMicroList.Clear
    For i = 1 To Me.sprQryList.DataRowCnt
        sprQryList.Row = i
        sprQryList.Col = 1: sLabno = sprQryList.Text
        sprQryList.Col = 2:
        frmResult.lstMicroList.AddItem Format(sLabno, "00000") & "  " & Trim(sprQryList.Text)
    Next
    Unload Me
    
    Return
    
    
End Sub

Private Sub txtPtno_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        If Trim(txtPtno.Text) = "" Then Exit Sub
        txtPtno.Text = Format(txtPtno.Text, "00000000")
        cmdQuery.SetFocus
    End If

    
    
End Sub

Private Sub txtSname_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        cmdQuery.SetFocus
    End If
    
End Sub
