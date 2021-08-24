VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmReJupsu 
   Caption         =   "검체도착확인"
   ClientHeight    =   7350
   ClientLeft      =   180
   ClientTop       =   1320
   ClientWidth     =   11655
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
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   11655
   WindowState     =   2  '최대화
   Begin Threed.SSPanel SSPanel4 
      Height          =   6675
      Left            =   135
      TabIndex        =   2
      Top             =   225
      Width           =   11490
      _Version        =   65536
      _ExtentX        =   20267
      _ExtentY        =   11774
      _StockProps     =   15
      Caption         =   "검체번호 개별 도착 확인"
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
      Begin FPSpreadADO.fpSpread sprItem 
         Height          =   4695
         Left            =   6525
         TabIndex        =   17
         Top             =   1710
         Width           =   3840
         _Version        =   196608
         _ExtentX        =   6773
         _ExtentY        =   8281
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
         MaxCols         =   1
         MaxRows         =   200
         ScrollBars      =   2
         SpreadDesigner  =   "frmReJupsu.frx":0000
         UserResize      =   1
         Appearance      =   1
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   690
         Left            =   720
         TabIndex        =   3
         Top             =   945
         Width           =   9645
         _Version        =   65536
         _ExtentX        =   17013
         _ExtentY        =   1217
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
         Begin VB.TextBox txtSname 
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Text            =   "txtSname"
            Top             =   225
            Width           =   870
         End
         Begin VB.TextBox txtSex 
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   2655
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Text            =   "txtSex"
            Top             =   225
            Width           =   330
         End
         Begin VB.TextBox txtAgeYY 
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   2970
            Locked          =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            Text            =   "txtAgeYY"
            Top             =   225
            Width           =   375
         End
         Begin VB.TextBox txtBirthDay 
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   5985
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Text            =   "txtBirthDay"
            Top             =   225
            Width           =   870
         End
         Begin VB.TextBox txtDeptName 
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   3330
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            Text            =   "txtDeptname"
            Top             =   225
            Width           =   1095
         End
         Begin VB.TextBox txtRoom 
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   5310
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            Text            =   "txtRoom"
            Top             =   225
            Width           =   690
         End
         Begin VB.TextBox txtDrname 
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   4410
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            Text            =   "txtDrname"
            Top             =   225
            Width           =   915
         End
         Begin VB.TextBox txtPtno 
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   855
            Locked          =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            Text            =   "txtPtno"
            Top             =   225
            Width           =   960
         End
         Begin VB.Label Label2 
            Caption         =   "환자정보"
            Height          =   240
            Left            =   90
            TabIndex        =   12
            Top             =   225
            Width           =   735
         End
      End
      Begin FPSpreadADO.fpSpread sprConfirm 
         Height          =   2400
         Left            =   675
         TabIndex        =   13
         Top             =   1710
         Width           =   5685
         _Version        =   196608
         _ExtentX        =   10028
         _ExtentY        =   4233
         _StockProps     =   64
         BackColorStyle  =   1
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
         MaxCols         =   9
         MaxRows         =   50
         ScrollBars      =   2
         SpreadDesigner  =   "frmReJupsu.frx":0ADA
         Appearance      =   2
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   555
         Left            =   720
         TabIndex        =   14
         Top             =   360
         Width           =   6270
         _Version        =   65536
         _ExtentX        =   11060
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
            TabIndex        =   15
            Top             =   180
            Width           =   915
         End
      End
      Begin MSForms.CommandButton cmdClear 
         Height          =   555
         Left            =   8820
         TabIndex        =   16
         Top             =   360
         Width           =   1545
         Caption         =   "Clear[F1]"
         PicturePosition =   327683
         Size            =   "2725;979"
         Picture         =   "frmReJupsu.frx":146B
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdOk 
         Height          =   555
         Left            =   7065
         TabIndex        =   1
         Top             =   360
         Width           =   1770
         Caption         =   "검체확인[F4]"
         PicturePosition =   327683
         Size            =   "3122;979"
         Picture         =   "frmReJupsu.frx":2BFD
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmReJupsu"
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
    
    sprConfirm.ReDraw = False
    sprConfirm.Row = 1
    sprConfirm.Row2 = sprConfirm.DataRowCnt
    sprConfirm.Col = 1
    sprConfirm.Col2 = sprConfirm.MaxCols
    sprConfirm.BlockMode = True
    sprConfirm.Action = ActionClearText
    sprConfirm.BlockMode = False
    sprConfirm.ReDraw = False
    
    
    sprItem.ReDraw = False
    sprItem.MaxRows = 0
    sprItem.MaxRows = 100
    sprItem.RowHeight(-1) = 11
    sprItem.ReDraw = True
    
    
End Sub

Private Sub cmdOk_Click()
    Dim sRowID      As String
    Dim nOrderno    As String
    Dim sOrderno    As String
    Dim iMatchno    As Integer
    
    
    
    
    If sprConfirm.DataRowCnt = 0 Then
        MsgBox "해당 접수 Data 가 하나도 없습니다!.. 접수한 Data를 먼저확인하세요"
        Exit Sub
    End If
    
    For i = 1 To sprConfirm.DataRowCnt
        sprConfirm.Row = i
        sprConfirm.Col = 1: sRowID = sprConfirm.Text
        sprConfirm.Col = 8: sOrderno = sprConfirm.Text
        sprConfirm.Col = 9: iMatchno = Val(sprConfirm.Text)
        GoSub Update_General_GbCH
    Next
    
    Call cmdClear_Click
    txtBarCode.SetFocus
    
    Exit Sub
    

Update_General_GbCH:
    Dim sEnrolTime  As String
    
    sEnrolTime = Dual_Date_Get("yyyy-MM-dd hh24:mi")
    
    strSql = ""
    strSql = strSql & " UPDATE TWEXAM_General"
    strSql = strSql & " SET    GbCH  = 'Y',"
    strSql = strSql & "        GBDate = TO_DATE('" & sEnrolTime & "','yyyy-MM-dd hh24:mi')"
    strSql = strSql & " WHERE  RowID = '" & sRowID & "'"
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    
  'Exam_Order Update
    Dim sJDt          As String
    Dim sSLno1        As String
    Dim sToDate       As String
    Dim nToHH         As Integer
    Dim nToMM         As Integer
    
    
    'sJDt = Left(txtBarCode.Text, 8)
    
    sJDt = convLabnoToExpand(Left(txtBarCode.Text, 5))
    
    sSLno1 = Val(Mid(txtBarCode.Text, 6, 2))

    sToDate = Dual_Date_Get("yyyy-MM-dd")
    nToHH = Val(Dual_Date_Get("hh"))
    nToMM = Val(Dual_Date_Get("mi"))
    
    strSql = ""
    strSql = strSql & " UPDATE TW_MIS_EXAM.TWEXAM_Order"
    strSql = strSql & " SET    GbCH      =  'Y',"
    strSql = strSql & "        GBDate    =   TO_DATE('" & sEnrolTime & "','yyyy-MM-dd hh24:mi'),"
    'strSql = strSql & "        CoLLDate  =   TO_DATE('" & sToDate & "','YYYY-MM-DD'),"
    'strSql = strSql & "        CoLLHH    =   " & nToHH & ","
    'strSql = strSql & "        CoLLMM    =   " & nToMM & ","
    strSql = strSql & "        CoLLid    =   " & Val(GstrIdnumber)
    strSql = strSql & " WHERE  JeobsuDt  =   TO_DATE('" & sJDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    SLipno1   =  '" & Val(sSLno1) & "'"
    'strSql = strSql & " AND    Orderno   =   " & Val(sOrderno) & ""
    strSql = strSql & " AND    Matchno  =   " & iMatchno & ""
    
    
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    Return
    

End Sub

Private Sub Form_Load()
    For i = 0 To Me.Count - 1
        If TypeOf Me.Controls(i) Is VB.TextBox Then
            Me.Controls(i).Text = ""
        End If
    Next
    
    sprConfirm.ReDraw = False
    sprConfirm.Row = 1
    sprConfirm.Row2 = sprConfirm.DataRowCnt
    sprConfirm.Col = 1
    sprConfirm.Col2 = sprConfirm.MaxCols
    sprConfirm.BlockMode = True
    sprConfirm.Action = ActionClearText
    sprConfirm.BlockMode = False
    sprConfirm.ReDraw = False
    
    
    sprItem.ReDraw = False
    sprItem.MaxRows = 0
    sprItem.MaxRows = 100
    sprItem.RowHeight(-1) = 11
    sprItem.ReDraw = True

End Sub

Private Sub mnuExit_Click()
    
    Unload Me
    
End Sub

Private Sub txtBarCode_GotFocus()
    
    txtBarCode.SelStart = 0
    txtBarCode.SelLength = Len(txtBarCode.Text)
    
End Sub

Private Sub txtBarCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sJeobsuDt        As String
    Dim iSLipno1         As Integer
    Dim iSLipno2         As Integer
    
    
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
        If Trim(txtPtno.Text) <> "" Then
            Call txtPtno_KeyDown(vbKeyReturn, 1)
        End If
        If sprConfirm.DataRowCnt > 0 Then
            cmdOk.SetFocus
        Else
            txtBarCode.SetFocus
            txtBarCode.SelStart = 0
            txtBarCode.SelLength = Len(txtBarCode.Text)
        End If
        
    End If
    
    Exit Sub
    

Get_General_Data:
    strSql = ""
    strSql = strSql & " SELECT a.RowID RwID, a.*,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDt,'YYYY-MM-DD') JeobsuDt,"
    strSql = strSql & "        b.Codenm SLipName, c.Codenm SampleName, a.Orderno"
    strSql = strSql & " FROM   TWEXAM_General a,"
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Specode b,"
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Sample  c "
    strSql = strSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyyMMdd')"
    strSql = strSql & " AND    a.SLipno1  = " & iSLipno1
    strSql = strSql & " AND    a.SLipno2  = " & iSLipno2
    strSql = strSql & " AND    a.GBCH    IN  ('1','2')"  '1=병동에서 접수, 2=정규채혈
    strSql = strSql & " AND    a.SLipno1  =  TO_Number(b.Codeky)"
    strSql = strSql & " AND    b.Codegu   =   '12'"
    strSql = strSql & " AND    a.GeomchCd =  c.Code(+)"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    Do Until adoSet.EOF
        sprConfirm.Row = sprConfirm.DataRowCnt + 1
        sprConfirm.Col = 1: sprConfirm.Text = adoSet.Fields("RwID").Value & ""
        sprConfirm.Col = 2: sprConfirm.Text = adoSet.Fields("JeobsuDt").Value & ""
        sprConfirm.Col = 3: sprConfirm.Text = adoSet.Fields("SampleName").Value & ""
        sprConfirm.Col = 4: sprConfirm.Text = adoSet.Fields("SLipName").Value & ""
        sprConfirm.Col = 5: sprConfirm.Text = adoSet.Fields("SLipno2").Value & ""
        sprConfirm.Col = 6: sprConfirm.Text = adoSet.Fields("GbCh").Value & ""
        
        Select Case adoSet.Fields("GbCh").Value & ""
            Case "1": sprConfirm.Col = 7: sprConfirm.Text = "병동 Or ER"
            Case "2": sprConfirm.Col = 7: sprConfirm.Text = "정규채혈"
        End Select
        
        txtPtno.Text = adoSet.Fields("Ptno").Value & ""
        sprConfirm.Col = 8: sprConfirm.Text = adoSet.Fields("Orderno").Value & ""
        sprConfirm.Col = 9: sprConfirm.Text = adoSet.Fields("Matchno").Value & ""
        adoSet.MoveNext
        
    Loop
    Call adoSetClose(adoSet)
    GoSub Select_ITemList
        
    Return
    
Select_ITemList:
    strSql = ""
    strSql = strSql & " SELECT DISTINCT b.Routinnm ItemName"
    strSql = strSql & " FROM   TWEXAM_General_Sub a,"
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Routine     b"
    strSql = strSql & " WHERE  a.JeobsuDt  = TO_DATE('" & sJeobsuDt & "','yyyyMMdd')"
    strSql = strSql & " AND    a.SLipno1   = " & iSLipno1
    strSql = strSql & " AND    a.SLipno2   = " & iSLipno2
    strSql = strSql & " AND    a.Routincd != a.ItemCd"
    strSql = strSql & " AND    a.Routincd  = b.Routincd(+)"
    strSql = strSql & " Union all"
    strSql = strSql & " SELECT b.ItemNM ItemName"
    strSql = strSql & " FROM   TWEXAM_General_Sub a,"
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_itemML      b"
    strSql = strSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyyMMdd')"
    strSql = strSql & " AND    a.SLipno1  = " & iSLipno1
    strSql = strSql & " AND    a.SLipno2  = " & iSLipno2
    strSql = strSql & " AND    a.Routincd = a.ItemCd"
    strSql = strSql & " AND    a.Itemcd   = b.Codeky(+)"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    Do Until adoSet.EOF
        sprItem.Row = sprItem.DataRowCnt + 1
        sprItem.Col = 1: sprItem.Text = adoSet.Fields("ItemName").Value & ""
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
        txtRoom.Text = adoSet.Fields("RoomCode").Value & ""
        txtSex.Text = adoSet.Fields("Sex").Value & ""
        txtAgeYY.Text = adoSet.Fields("AgeYY").Value & ""
        txtBirthDay.Text = adoSet.Fields("BirthDay").Value & ""
        txtDeptName.Text = adoSet.Fields("DeptnameK").Value & ""
        txtDrname.Text = adoSet.Fields("Drname").Value & ""
        Call adoSetClose(adoSet)
    End If
        
    Return

End Sub
