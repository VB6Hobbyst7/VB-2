VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{48932A52-981F-101B-A7FB-4A79242FD97B}#2.0#0"; "TAB32X20.OCX"
Begin VB.Form frmWhere 
   Caption         =   "조회조건화면"
   ClientHeight    =   2805
   ClientLeft      =   2610
   ClientTop       =   2775
   ClientWidth     =   5880
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   5880
   Begin Threed.SSPanel SSPanel1 
      Height          =   510
      Left            =   90
      TabIndex        =   12
      Top             =   45
      Width           =   5685
      _Version        =   65536
      _ExtentX        =   10028
      _ExtentY        =   900
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComCtl2.DTPicker dtToDate 
         Height          =   330
         Left            =   3150
         TabIndex        =   14
         Top             =   90
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24379395
         CurrentDate     =   36379
      End
      Begin MSComCtl2.DTPicker dtFrDate 
         Height          =   330
         Left            =   1620
         TabIndex        =   13
         Top             =   90
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24379395
         CurrentDate     =   36379
      End
      Begin VB.Label Label5 
         Caption         =   "Date:From/To"
         Height          =   240
         Left            =   360
         TabIndex        =   16
         Top             =   135
         Width           =   1140
      End
   End
   Begin TabproLib.vaTabPro vaTabPro2 
      Height          =   2175
      Left            =   90
      TabIndex        =   0
      Top             =   585
      Width           =   5730
      _Version        =   131072
      _ExtentX        =   10107
      _ExtentY        =   3836
      _StockProps     =   100
      AlignTextV      =   1
      Orientation     =   2
      TabShape        =   3
      ApplyTo         =   2
      OffsetFromClientTop=   -1  'True
      ChamferedWidth  =   1
      ChamferedHeight =   1
      Mode            =   1
      BookCornerType  =   1
      BookRingShowHole=   -1  'True
      BookShowCornerGuard=   -1  'True
      BookCornerGuardWidth=   105
      BookCornerGuardLength=   405
      TabCaption      =   "frmWhere.frx":0000
      Begin Threed.SSCommand cmdComboCls 
         Height          =   285
         Index           =   0
         Left            =   2925
         TabIndex        =   17
         Top             =   675
         Width           =   195
         _Version        =   65536
         _ExtentX        =   344
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "C"
      End
      Begin VB.TextBox txtQryptno 
         Enabled         =   0   'False
         Height          =   330
         Left            =   -17639
         MaxLength       =   8
         TabIndex        =   9
         Top             =   -15629
         Width           =   1455
      End
      Begin VB.TextBox txtQrysname 
         Enabled         =   0   'False
         Height          =   330
         Left            =   -17549
         TabIndex        =   6
         Top             =   -15629
         Width           =   1365
      End
      Begin VB.ComboBox cmbRoom 
         Height          =   300
         Left            =   1845
         Style           =   2  '드롭다운 목록
         TabIndex        =   4
         Top             =   990
         Width           =   1050
      End
      Begin VB.ComboBox cmbWard 
         Height          =   300
         Left            =   855
         Style           =   2  '드롭다운 목록
         TabIndex        =   2
         Top             =   675
         Width           =   2040
      End
      Begin Threed.SSCommand cmdComboCls 
         Height          =   285
         Index           =   1
         Left            =   2925
         TabIndex        =   18
         Top             =   990
         Width           =   195
         _Version        =   65536
         _ExtentX        =   344
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "C"
      End
      Begin MSForms.CommandButton cmdQry3 
         Height          =   600
         Left            =   -19934
         TabIndex        =   15
         Top             =   -15899
         Width           =   1545
         VariousPropertyBits=   25
         Caption         =   "조회확인"
         PicturePosition =   327683
         Size            =   "2725;1058"
         Picture         =   "frmWhere.frx":03BE
         FontName        =   "굴림체"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdQry2 
         Height          =   600
         Left            =   -19934
         TabIndex        =   11
         Top             =   -15899
         Width           =   1545
         VariousPropertyBits=   25
         Caption         =   "조회확인"
         PicturePosition =   327683
         Size            =   "2725;1058"
         Picture         =   "frmWhere.frx":0C98
         FontName        =   "굴림체"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin VB.Label Label3 
         Caption         =   "병록번호"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -16064
         TabIndex        =   10
         Top             =   -15674
         Width           =   735
      End
      Begin MSForms.CommandButton cmdQry1 
         Height          =   600
         Left            =   -19934
         TabIndex        =   8
         Top             =   -15899
         Width           =   1545
         VariousPropertyBits=   25
         Caption         =   "조회확인"
         PicturePosition =   327683
         Size            =   "2725;1058"
         Picture         =   "frmWhere.frx":1572
         FontName        =   "굴림체"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin VB.Label Label2 
         Caption         =   "재원자명"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -16064
         TabIndex        =   7
         Top             =   -15674
         Width           =   735
      End
      Begin MSForms.CommandButton cmbQry0 
         Height          =   600
         Left            =   3510
         TabIndex        =   5
         Top             =   675
         Width           =   1545
         Caption         =   "조회확인"
         PicturePosition =   327683
         Size            =   "2725;1058"
         Picture         =   "frmWhere.frx":1E4C
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin VB.Label Label1 
         Caption         =   "병실"
         Height          =   240
         Left            =   1350
         TabIndex        =   3
         Top             =   1035
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "병동"
         Height          =   240
         Left            =   405
         TabIndex        =   1
         Top             =   720
         Width           =   510
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmWhere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sFrJeobsuDt     As String
Dim sToJeobsuDt     As String
Public Function Spread_Set_Clear(ByVal sSpread As Object) As Integer
    
    sSpread.Row = 1
    sSpread.Row2 = sSpread.DataRowCnt
    sSpread.Col = 1
    sSpread.Col2 = sSpread.DataColCnt
    sSpread.BlockMode = True
    sSpread.Action = ActionClear
    sSpread.BlockMode = False
    

End Function


Private Sub cmbQry0_Click()
    
    
    sFrJeobsuDt = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToJeobsuDt = Format(dtToDate.Value, "yyyy-MM-dd")
    
    If Trim(cmbRoom.Text) <> "" Then  '병실까지 선택했을경우
        StrSql = ""
        StrSql = StrSql & " SELECT TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
        StrSql = StrSql & "        a.Ptno, a.Roomcode, b.Sname"
        StrSql = StrSql & " FROM   TWEXAM_General a,"
        StrSql = StrSql & "        TWBAS_Patient  b "
        StrSql = StrSql & " WHERE  a.JeobsuDt >=      TO_DATE('" & sFrJeobsuDt & "','YYYY-MM-DD')"
        StrSql = StrSql & " AND    a.JeobsuDt <=      TO_DATE('" & sToJeobsuDt & "','YYYY-MM-DD')"
        StrSql = StrSql & " AND    a.RoomCode  = '" & cmbRoom.Text & "'"
        StrSql = StrSql & " AND    a.Ptno      =  b.Ptno(+)"
        StrSql = StrSql & " GROUP  BY a.Ptno, JeobsuDt, RoomCode, Sname"
    ElseIf Trim(cmbWard.Text) <> "" Then    '병동만 선택
        StrSql = ""
        StrSql = StrSql & " SELECT TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
        StrSql = StrSql & "         a.Ptno, a.Roomcode, b.Sname"
        StrSql = StrSql & " FROM   TWEXAM_General a,"
        StrSql = StrSql & "        TWBAS_Patient  b,"
        StrSql = StrSql & "        TWBAS_ROOM     c "
        StrSql = StrSql & " WHERE  a.JeobsuDt >=    TO_DATE('" & sFrJeobsuDt & "','YYYY-MM-DD')"
        StrSql = StrSql & " AND    a.JeobsuDt <=    TO_DATE('" & sToJeobsuDt & "','YYYY-MM-DD')"
        StrSql = StrSql & " AND    a.Ptno      =  b.Ptno(+)"
        StrSql = StrSql & " AND    a.RoomCode  =  c.RoomCode"
        StrSql = StrSql & " AND    c.WardCode  =  '" & Left(cmbWard.Text, 4) & "'"
        StrSql = StrSql & " GROUP  BY a.RoomCode, JeobsuDt, a.Ptno,Sname"
    Else                               '아무것도 선택하지 않음
        StrSql = ""
        StrSql = StrSql & " SELECT TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
        StrSql = StrSql & "        a.Ptno, a.Roomcode, b.Sname"
        StrSql = StrSql & " FROM   TWEXAM_General a,"
        StrSql = StrSql & "        TWBAS_Patient  b "
        StrSql = StrSql & " WHERE  a.JeobsuDt >=      TO_DATE('" & sFrJeobsuDt & "','YYYY-MM-DD')"
        StrSql = StrSql & " AND    a.JeobsuDt <=      TO_DATE('" & sToJeobsuDt & "','YYYY-MM-DD')"
        StrSql = StrSql & " AND    a.Ptno      =  b.Ptno(+)"
        StrSql = StrSql & " GROUP  BY RoomCode, JeobsuDt, a.Ptno, Sname"
    End If
        
    
    If False = adoSetOpen(StrSql, adoSet) Then Exit Sub
    Call Spread_Set_Clear(frmIpdLabel.ssPtList)
    
    Do Until adoSet.EOF
        frmIpdLabel.ssPtList.Row = frmIpdLabel.ssPtList.DataRowCnt + 1
        frmIpdLabel.ssPtList.Col = 2: frmIpdLabel.ssPtList.Text = adoSet.Fields("Roomcode").Value & ""
        frmIpdLabel.ssPtList.Col = 3: frmIpdLabel.ssPtList.Text = adoSet.Fields("Ptno").Value & ""
        frmIpdLabel.ssPtList.Col = 4: frmIpdLabel.ssPtList.Text = adoSet.Fields("Sname").Value & ""
        frmIpdLabel.ssPtList.Col = 5: frmIpdLabel.ssPtList.Text = adoSet.Fields("JeobsuDt").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
End Sub

Private Sub cmbWard_Click()
    If cmbWard.ListIndex = -1 Then Exit Sub
    
    StrSql = ""
    StrSql = StrSql & " SELECT RoomCode"
    StrSql = StrSql & " FROM   TWBAS_Room"
    StrSql = StrSql & " WHERE  WardCode = '" & Left(cmbWard.Text, 4) & "'"
    If False = adoSetOpen(StrSql, adoSet) Then Return
    
    cmbRoom.Clear
    Do Until adoSet.EOF
        cmbRoom.AddItem adoSet.Fields("RoomCode").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    
End Sub

Private Sub cmdComboCls_Click(Index As Integer)
    
    If Index = 0 Then
        cmbWard.ListIndex = -1
        cmbRoom.ListIndex = -1
    Else
        cmbRoom.ListIndex = -1
    End If
    
End Sub

Private Sub cmdQry1_Click()
    
    sFrJeobsuDt = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToJeobsuDt = Format(dtToDate.Value, "yyyy-MM-dd")
    
    StrSql = ""
    StrSql = StrSql & " SELECT TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
    StrSql = StrSql & "        a.Ptno, a.Roomcode, b.Sname"
    StrSql = StrSql & " FROM   TWEXAM_General a,"
    StrSql = StrSql & "        TWIPD_Master   b "
    StrSql = StrSql & " WHERE  a.JeobsuDt >=      TO_DATE('" & sFrJeobsuDt & "','YYYY-MM-DD')"
    StrSql = StrSql & " AND    a.JeobsuDt <=      TO_DATE('" & sToJeobsuDt & "','YYYY-MM-DD')"
    StrSql = StrSql & " AND    b.Sname     Like '" & txtQrysname.Text & "%'"
    StrSql = StrSql & " AND    a.Ptno      =  b.Ptno(+)"
    StrSql = StrSql & " GROUP  BY JeobsuDt, a.Ptno, a.RoomCode, b.Sname"
        
    
    If False = adoSetOpen(StrSql, adoSet) Then Exit Sub
    Call Spread_Set_Clear(frmIpdLabel.ssPtList)
    
    Do Until adoSet.EOF
        frmIpdLabel.ssPtList.Row = frmIpdLabel.ssPtList.DataRowCnt + 1
        frmIpdLabel.ssPtList.Col = 2: frmIpdLabel.ssPtList.Text = adoSet.Fields("Roomcode").Value & ""
        frmIpdLabel.ssPtList.Col = 3: frmIpdLabel.ssPtList.Text = adoSet.Fields("Ptno").Value & ""
        frmIpdLabel.ssPtList.Col = 4: frmIpdLabel.ssPtList.Text = adoSet.Fields("Sname").Value & ""
        frmIpdLabel.ssPtList.Col = 5: frmIpdLabel.ssPtList.Text = adoSet.Fields("JeobsuDt").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
End Sub

Private Sub cmdQry2_Click()
    
    sFrJeobsuDt = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToJeobsuDt = Format(dtToDate.Value, "yyyy-MM-dd")
    
    Call cmdClear_Click
    StrSql = ""
    StrSql = StrSql & " SELECT TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
    StrSql = StrSql & "        a.Ptno, a.Roomcode, b.Sname"
    StrSql = StrSql & " FROM   TWEXAM_General a,"
    StrSql = StrSql & "        TWIPD_Master   b "
    StrSql = StrSql & " WHERE  a.JeobsuDt >=      TO_DATE('" & sFrJeobsuDt & "','YYYY-MM-DD')"
    StrSql = StrSql & " AND    a.JeobsuDt <=      TO_DATE('" & sToJeobsuDt & "','YYYY-MM-DD')"
    StrSql = StrSql & " AND    b.Ptno      =  '" & txtQryptno.Text & "'"
    StrSql = StrSql & " AND    a.Ptno      =  b.Ptno(+)"
    StrSql = StrSql & " GROUP  BY JeobsuDt, a.Ptno, a.RoomCode, b.Sname"
        
    If False = adoSetOpen(StrSql, adoSet) Then Exit Sub
    Call Spread_Set_Clear(frmIpdLabel.ssPtList)
    
    Do Until adoSet.EOF
        frmIpdLabel.ssPtList.Row = frmIpdLabel.ssPtList.DataRowCnt + 1
        frmIpdLabel.ssPtList.Col = 2: frmIpdLabel.ssPtList.Text = adoSet.Fields("Roomcode").Value & ""
        frmIpdLabel.ssPtList.Col = 3: frmIpdLabel.ssPtList.Text = adoSet.Fields("Ptno").Value & ""
        frmIpdLabel.ssPtList.Col = 4: frmIpdLabel.ssPtList.Text = adoSet.Fields("Sname").Value & ""
        frmIpdLabel.ssPtList.Col = 5: frmIpdLabel.ssPtList.Text = adoSet.Fields("JeobsuDt").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)

End Sub

Private Sub cmdQry3_Click()
    
    sFrJeobsuDt = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToJeobsuDt = Format(dtToDate.Value, "yyyy-MM-dd")
    
    Call cmdClear_Click
    
    StrSql = ""
    StrSql = StrSql & " SELECT TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
    StrSql = StrSql & "        a.Ptno, a.Roomcode, b.Sname"
    StrSql = StrSql & " FROM   TWEXAM_General a,"
    StrSql = StrSql & "        TWIPD_Master   b "
    StrSql = StrSql & " WHERE  a.JeobsuDt >=      TO_DATE('" & sFrJeobsuDt & "','YYYY-MM-DD')"
    StrSql = StrSql & " AND    a.JeobsuDt <=      TO_DATE('" & sToJeobsuDt & "','YYYY-MM-DD')"
    StrSql = StrSql & " AND    a.Ptno      =  b.Ptno(+)"
    StrSql = StrSql & " GROUP  BY JeobsuDt, a.Ptno, a.RoomCode, b.Sname"
        
    If False = adoSetOpen(StrSql, adoSet) Then Exit Sub
    Call Spread_Set_Clear(frmIpdLabel.ssPtList)
    
    Do Until adoSet.EOF
        frmIpdLabel.ssPtList.Row = frmIpdLabel.ssPtList.DataRowCnt + 1
        frmIpdLabel.ssPtList.Col = 2: frmIpdLabel.ssPtList.Text = adoSet.Fields("Roomcode").Value & ""
        frmIpdLabel.ssPtList.Col = 3: frmIpdLabel.ssPtList.Text = adoSet.Fields("Ptno").Value & ""
        frmIpdLabel.ssPtList.Col = 4: frmIpdLabel.ssPtList.Text = adoSet.Fields("Sname").Value & ""
        frmIpdLabel.ssPtList.Col = 5: frmIpdLabel.ssPtList.Text = adoSet.Fields("JeobsuDt").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)

End Sub

Private Sub Form_Load()
    
    GoSub Form_Center
    GoSub Get_Dual_SysDate
    GoSub Get_Ward_Data
    Exit Sub



Form_Center:
    Me.Height = 3500
    Me.Width = 6000
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Return
    
Get_Dual_SysDate:
    dtFrDate.Value = Dual_Date_Get("yyyy-MM-dd")
    dtToDate.Value = Dual_Date_Get("yyyy-MM-dd")
    Return
    
Get_Ward_Data:
    Dim sWardC  As String * 4
    
    StrSql = ""
    StrSql = StrSql & " SELECT *"
    StrSql = StrSql & " FROM   TWBAS_WARD"
    StrSql = StrSql & " ORDER  BY WardCode"
    If False = adoSetOpen(StrSql, adoSet) Then Return
    Do Until adoSet.EOF
        sWardC = adoSet.Fields("WardCode").Value & ""
        cmbWard.AddItem sWardC & Trim(adoSet.Fields("WardName").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return


End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

Private Sub Text2_Change()

End Sub

Private Sub txtQryptno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdQry2.SetFocus
    End If

End Sub

Private Sub txtQryptno_LostFocus()
    
    txtQryptno.Text = UCase(txtQryptno.Text)
    txtQryptno.Text = Format(txtQryptno.Text, "00000000")
    
End Sub

Private Sub txtQrysname_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        cmdQry1.SetFocus
    End If
End Sub
