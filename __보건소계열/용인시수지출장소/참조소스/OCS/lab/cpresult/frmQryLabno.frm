VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmQryLabno 
   Caption         =   "Labno로 검사를 조회합니다!..."
   ClientHeight    =   7020
   ClientLeft      =   120
   ClientTop       =   1410
   ClientWidth     =   11595
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7020
   ScaleWidth      =   11595
   WindowState     =   2  '최대화
   Begin FPSpreadADO.fpSpread sprGeneral 
      Height          =   4110
      Left            =   540
      TabIndex        =   22
      Top             =   2295
      Width           =   9645
      _Version        =   196608
      _ExtentX        =   17013
      _ExtentY        =   7250
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
      MaxRows         =   100
      ScrollBars      =   2
      SpreadDesigner  =   "frmQryLabno.frx":0000
      UserResize      =   0
      Appearance      =   2
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   405
      Top             =   5355
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQryLabno.frx":0797
            Key             =   "Exit"
            Object.Tag             =   "Exit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '위 맞춤
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   635
      ButtonWidth     =   1270
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit of Screen"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   1140
      Left            =   540
      TabIndex        =   2
      Top             =   1170
      Width           =   9645
      _Version        =   65536
      _ExtentX        =   17013
      _ExtentY        =   2011
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
      BorderWidth     =   0
      BevelInner      =   1
      Alignment       =   0
      Begin VB.TextBox txtStatus 
         BackColor       =   &H00C00000&
         ForeColor       =   &H80000009&
         Height          =   330
         Left            =   7380
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   495
         Width           =   1545
      End
      Begin VB.TextBox txtPtno 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   855
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "txtPtno"
         Top             =   225
         Width           =   1140
      End
      Begin VB.TextBox txtDrname 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "txtDrname"
         Top             =   525
         Width           =   1140
      End
      Begin VB.TextBox txtRoom 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "txtRoom"
         Top             =   195
         Width           =   1140
      End
      Begin VB.TextBox txtDeptName 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   2970
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "txtDeptname"
         Top             =   525
         Width           =   1140
      End
      Begin VB.TextBox txtBirthDay 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   7875
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "txtBirthDay"
         Top             =   195
         Width           =   1050
      End
      Begin VB.TextBox txtAgeYY 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   7380
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "txtAgeYY"
         Top             =   195
         Width           =   465
      End
      Begin VB.TextBox txtSex 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "txtSex"
         Top             =   195
         Width           =   375
      End
      Begin VB.TextBox txtSname 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   2970
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "txtSname"
         Top             =   195
         Width           =   1140
      End
      Begin VB.Label Label2 
         Caption         =   "등록번호"
         Height          =   240
         Left            =   90
         TabIndex        =   17
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "환자명"
         Height          =   195
         Left            =   2205
         TabIndex        =   16
         Top             =   270
         Width           =   690
      End
      Begin VB.Label Label4 
         Caption         =   "병실"
         Height          =   195
         Left            =   4275
         TabIndex        =   15
         Top             =   225
         Width           =   420
      End
      Begin VB.Label Label5 
         Caption         =   "성별"
         Height          =   195
         Left            =   6030
         TabIndex        =   14
         Top             =   225
         Width           =   420
      End
      Begin VB.Label Label6 
         Caption         =   "나이"
         Height          =   240
         Left            =   6975
         TabIndex        =   13
         Top             =   225
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "과"
         Height          =   195
         Left            =   2205
         TabIndex        =   12
         Top             =   585
         Width           =   600
      End
      Begin VB.Label Label8 
         Caption         =   "의사"
         Height          =   240
         Left            =   4275
         TabIndex        =   11
         Top             =   540
         Width           =   600
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   555
      Left            =   540
      TabIndex        =   18
      Top             =   540
      Width           =   4065
      _Version        =   65536
      _ExtentX        =   7170
      _ExtentY        =   979
      _StockProps     =   15
      BackColor       =   8421376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      Begin VB.TextBox txtBarCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1215
         TabIndex        =   0
         Top             =   135
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808000&
         Caption         =   "BarCode :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   135
         TabIndex        =   19
         Top             =   180
         Width           =   915
      End
   End
   Begin MSForms.CommandButton cmdClear 
      Height          =   555
      Left            =   4680
      TabIndex        =   20
      Top             =   540
      Width           =   1905
      Caption         =   "   Clear[F1]"
      PicturePosition =   327683
      Size            =   "3360;979"
      Picture         =   "frmQryLabno.frx":0AB3
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmQryLabno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    
    For i = 0 To Me.Count - 1
        If TypeOf Me.Controls(i) Is VB.TextBox Then
            Me.Controls(i).Text = ""
        End If
    Next
    
    sprGeneral.Row = 1
    sprGeneral.Row2 = sprGeneral.DataRowCnt
    sprGeneral.Col = 1
    sprGeneral.Col2 = sprGeneral.MaxCols
    sprGeneral.BlockMode = True
    sprGeneral.Action = ActionClearText
    sprGeneral.BlockMode = False
    txtBarCode.SetFocus
    

End Sub

Private Sub cmdOk_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF1 Then
        Call cmdClear_Click
    End If
    
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1: Unload Me
        
    End Select
    
End Sub

Private Sub txtBarCode_GotFocus()
    
    txtBarCode.SelStart = 0
    
    txtBarCode.SelLength = Len(txtBarCode.Text)
    
End Sub

Private Sub txtBarCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sJeobsuDt        As String
    Dim iSLipno1         As Integer
    Dim iSLipno2         As Integer
    Dim sKey             As String
    
    
    If KeyCode = vbKeyReturn Then
        If Trim(txtBarCode.Text) = "" Then Exit Sub
        txtBarCode.Tag = txtBarCode.Text
        Call cmdClear_Click
        txtBarCode.Text = txtBarCode.Tag
        
        Select Case Len(Trim(txtBarCode.Text))
            Case 12
                sJeobsuDt = convLabnoToExpand(Left(txtBarCode.Text, 5))
                iSLipno1 = Val(Mid(txtBarCode.Text, 6, 2))
                iSLipno2 = Val(Mid(txtBarCode.Text, 8, 5))
            Case 15
                sJeobsuDt = Left(txtBarCode.Text, 8)
                iSLipno1 = Val(Mid(txtBarCode.Text, 9, 2))
                iSLipno2 = Val(Mid(txtBarCode.Text, 11, 5))
        End Select
        
        GoSub Get_General_Data
        GoSub Get_Patient_Data
        txtBarCode.SelStart = 0
        txtBarCode.SelLength = Len(txtBarCode.Text)
    End If
    
    Exit Sub
    
Get_Patient_Data:
    strSql = ""
    strSql = strSql & " SELECT a.*, b.Sname, b.BirthDay, c.DeptNamek, d.Drname, "
    strSql = strSql & "        e.RoomCode IPDRoom"
    strSql = strSql & " FROM   TWEXAM_General a,"
    strSql = strSql & "        TWEXAM_IDNOMST b,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT     c,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR   d,"
    strSql = strSql & "        TW_MIS_PMPA.TWIPD_MASTER   e"
    strSql = strSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyyMMdd')"
    strSql = strSql & " AND    a.SLipno1  = " & iSLipno1
    strSql = strSql & " AND    a.SLipno2  = " & iSLipno2
    strSql = strSql & " AND    a.Ptno     = b.Ptno(+)"
    strSql = strSql & " AND    a.DeptCode = c.DeptCode(+)"
    strSql = strSql & " AND    a.DrCode   = d.Drcode(+)"
    strSql = strSql & " AND    a.Ptno     = e.Ptno(+)"
    strSql = strSql & " AND    e.amSet6(+)   = ' '"
    
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    txtPtno.Text = adoSet.Fields("Ptno").Value & ""
    txtSname.Text = adoSet.Fields("Sname").Value & ""
    txtSex.Text = adoSet.Fields("Sex").Value & ""
    txtAgeYY.Text = adoSet.Fields("AgeYY").Value & ""
    txtBirthDay.Text = adoSet.Fields("BirthDay").Value & ""
    txtDeptName.Text = adoSet.Fields("DeptNameK").Value & ""
    txtDrname.Text = adoSet.Fields("Drname").Value & ""
    txtRoom.Text = adoSet.Fields("IPDRoom").Value & ""
    Call adoSetClose(adoSet)
    

    Return
    

Get_General_Data:
    strSql = ""
    strSql = strSql & " SELECT a.Sex, a.AgeYY, a.Deptcode, a.Drcode, a.Roomcode, a.Status,"
    strSql = strSql & "        TO_CHAR(b.JeobsuDt, 'yyyy-MM-dd') JeobsuDt,"
    strSql = strSql & "        b.SLipno1, b.SLipno2,"
    strSql = strSql & "        b.ItemCd ItemCode, c.ItemNM ItemName,"
    strSql = strSql & "        d.Codenm SampleName"
    strSql = strSql & " FROM   TWEXAM_General      a,"
    strSql = strSql & "        TWEXAM_General_Sub  b,"
    strSql = strSql & "        TWEXAM_ItemML       c,"
    strSql = strSql & "        TWEXAM_Sample       d "
    strSql = strSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyyMMdd')"
    strSql = strSql & " AND    a.SLipno1  = " & iSLipno1
    strSql = strSql & " AND    a.SLipno2  = " & iSLipno2
    strSql = strSql & " AND    a.JeobsuDt = b.JeobsuDt(+)"
    strSql = strSql & " AND    a.SLipno1  = b.SLipno1(+)"
    strSql = strSql & " AND    a.SLipno2  = b.SLipno2(+)"
    strSql = strSql & " AND    b.RoutinCD = b.ItemCd"      'Item Order
    strSql = strSql & " AND    b.ItemCd   = c.Codeky(+)"
    strSql = strSql & " AND    a.GeomchCd = d.Code(+)"
    strSql = strSql & " UNION ALL"
    strSql = strSql & " SELECT DISTINCT a.Sex, a.AgeYY, a.Deptcode, a.Drcode, a.Roomcode,a.Status,"
    strSql = strSql & "        TO_CHAR(b.JeobsuDt, 'yyyy-MM-dd') JeobsuDt,"
    strSql = strSql & "        b.SLipno1, b.SLipno2,"
    strSql = strSql & "        b.RoutinCd ItemCode , c.RoutinNM ItemName,"
    strSql = strSql & "        d.Codenm SampleName"
    strSql = strSql & " FROM   TWEXAM_General      a,"
    strSql = strSql & "        TWEXAM_General_Sub  b,"
    strSql = strSql & "        TWEXAM_Routine      c,"
    strSql = strSql & "        TWEXAM_Sample       d "
    strSql = strSql & " WHERE  a.JeobsuDt  = TO_DATE('" & sJeobsuDt & "','yyyyMMdd')"
    strSql = strSql & " AND    a.SLipno1   = " & iSLipno1
    strSql = strSql & " AND    a.SLipno2   = " & iSLipno2
    strSql = strSql & " AND    a.JeobsuDt  = b.JeobsuDt(+)"
    strSql = strSql & " AND    a.SLipno1   = b.SLipno1(+)"
    strSql = strSql & " AND    a.SLipno2   = b.SLipno2(+)"
    strSql = strSql & " AND    b.RoutinCD != b.ItemCd"      'Routine Order
    strSql = strSql & " AND    b.RoutinCd  = c.RoutinCd(+)"
    strSql = strSql & " AND    a.GeomchCd =  d.Code(+)"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    Do Until adoSet.EOF
        sprGeneral.Row = sprGeneral.DataRowCnt + 1
        If sKey <> adoSet.Fields("JeobsuDt").Value & "" & _
                   adoSet.Fields("SLipno1").Value & "" & _
                   adoSet.Fields("SLipno2").Value & "" Then
        
            sprGeneral.Col = 1: sprGeneral.Text = adoSet.Fields("JeobsuDt").Value & ""
            sprGeneral.Col = 2: sprGeneral.Text = adoSet.Fields("SLipno1").Value & ""
            sprGeneral.Col = 3: sprGeneral.Text = adoSet.Fields("SLipno2").Value & ""
        End If
        
        sprGeneral.Col = 4: sprGeneral.Text = adoSet.Fields("ItemCode").Value & ""
        sprGeneral.Col = 5: sprGeneral.Text = adoSet.Fields("ItemName").Value & ""
        sprGeneral.Col = 6: sprGeneral.Text = adoSet.Fields("SampleName").Value & ""
        
        txtStatus.Text = ""
        Select Case Trim(adoSet.Fields("Status").Value & "")
            Case "R": txtStatus.Text = "접수중"
            Case "P": txtStatus.Text = "부분결과"
            Case "U": txtStatus.Text = "미확인"
            Case "C": txtStatus.Text = "결과완료"
            Case "X": txtStatus.Text = "이상Data"
        End Select
        
       sKey = adoSet.Fields("JeobsuDt").Value & "" & _
              adoSet.Fields("SLipno1").Value & "" & _
              adoSet.Fields("SLipno2").Value & ""
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
        
    Return

End Sub

Private Sub txtPtno_KeyDown(KeyCode As Integer, Shift As Integer)
    
    
    If KeyCode = vbKeyReturn Then
        GoSub Check_Ptno_Text
        GoSub Main_Search_Data
    End If
    Exit Sub
    


Check_Ptno_Text:
    
    If Trim(txtPtno.Text) = "" Then Exit Sub
    If Not IsNumeric(txtPtno.Text) Then Exit Sub
    txtPtno.Text = Format(txtPtno.Text, "00000000")
        
    Return
    
    
Main_Search_Data:
    strSql = ""
    strSql = strSql & " SELECT b.WardCode, a.*, c.DeptnameK, d.Drname"
    strSql = strSql & " FROM   TWEXAM_IDNOMST a,"
    strSql = strSql & "        TW_MIS_PMPA.TWBas_Room     b,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT     c,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR   d "
    strSql = strSql & " WHERE  a.Ptno     = '" & txtPtno.Text & "'"
    strSql = strSql & " AND    a.RoomCode = b.RoomCode(+)"
    strSql = strSql & " AND    a.DeptCode = c.DeptCode(+)"
    strSql = strSql & " AND    a.DrCode   = d.DrCode(+)"
    
    
    If False = adoSetOpen(strSql, adoSet) Then
        MsgBox "등록번호 " & txtPtno.Text & " 는(은) 접수된 Data 가 없습니다!..."
        Call cmdClear_Click
        Exit Sub
    Else
        txtSname.Text = adoSet.Fields("Sname").Value & ""
        txtBirthDay.Text = adoSet.Fields("BirthDay").Value & ""
        Call adoSetClose(adoSet)
    End If
        
    Return

End Sub
