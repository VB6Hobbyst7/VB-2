VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmReEnrol2 
   Caption         =   "검체도착확인"
   ClientHeight    =   7350
   ClientLeft      =   315
   ClientTop       =   1380
   ClientWidth     =   11475
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
   ScaleHeight     =   7350
   ScaleWidth      =   11475
   Begin Threed.SSPanel SSPanel4 
      Height          =   6630
      Left            =   135
      TabIndex        =   0
      Top             =   270
      Width           =   11265
      _Version        =   65536
      _ExtentX        =   19870
      _ExtentY        =   11695
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   2625
         Left            =   495
         TabIndex        =   1
         Top             =   1080
         Width           =   2715
         _Version        =   65536
         _ExtentX        =   4789
         _ExtentY        =   4630
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
         BorderWidth     =   2
         Alignment       =   0
         Begin VB.TextBox txtSname 
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   855
            Locked          =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            Text            =   "txtSname"
            Top             =   555
            Width           =   1140
         End
         Begin VB.TextBox txtSex 
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   855
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Text            =   "txtSex"
            Top             =   1215
            Width           =   465
         End
         Begin VB.TextBox txtAgeYY 
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   855
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            Text            =   "txtAgeYY"
            Top             =   1545
            Width           =   465
         End
         Begin VB.TextBox txtBirthDay 
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   1350
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            Text            =   "txtBirthDay"
            Top             =   1545
            Width           =   1140
         End
         Begin VB.TextBox txtDeptName 
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   855
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            Text            =   "txtDeptname"
            Top             =   1875
            Width           =   1140
         End
         Begin VB.TextBox txtRoom 
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   855
            Locked          =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            Text            =   "txtRoom"
            Top             =   885
            Width           =   1140
         End
         Begin VB.TextBox txtDrname 
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   855
            Locked          =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            Text            =   "txtDrname"
            Top             =   2205
            Width           =   1140
         End
         Begin VB.TextBox txtPtno 
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   855
            Locked          =   -1  'True
            TabIndex        =   2
            TabStop         =   0   'False
            Text            =   "txtPtno"
            Top             =   225
            Width           =   1140
         End
         Begin VB.Label Label8 
            Caption         =   "의사"
            Height          =   240
            Left            =   90
            TabIndex        =   16
            Top             =   2250
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "과"
            Height          =   240
            Left            =   90
            TabIndex        =   15
            Top             =   1935
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "나이"
            Height          =   240
            Left            =   90
            TabIndex        =   14
            Top             =   1575
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "성별"
            Height          =   195
            Left            =   90
            TabIndex        =   13
            Top             =   1260
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "병실"
            Height          =   240
            Left            =   90
            TabIndex        =   12
            Top             =   900
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "환자명"
            Height          =   240
            Left            =   90
            TabIndex        =   11
            Top             =   585
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "등록번호"
            Height          =   240
            Left            =   90
            TabIndex        =   10
            Top             =   270
            Width           =   735
         End
      End
      Begin FPSpreadADO.fpSpread sprConfirm 
         Height          =   4200
         Left            =   3240
         TabIndex        =   17
         Top             =   1080
         Width           =   7080
         _Version        =   196608
         _ExtentX        =   12488
         _ExtentY        =   7408
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
         MaxCols         =   8
         MaxRows         =   50
         ScrollBars      =   2
         SpreadDesigner  =   "frmReEnrol2.frx":0000
         Appearance      =   2
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   555
         Left            =   495
         TabIndex        =   18
         Top             =   450
         Width           =   4695
         _Version        =   65536
         _ExtentX        =   8281
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
            TabIndex        =   19
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
            TabIndex        =   20
            Top             =   180
            Width           =   915
         End
      End
      Begin MSForms.CommandButton cmdClear 
         Height          =   555
         Left            =   7200
         TabIndex        =   22
         Top             =   450
         Width           =   1905
         Caption         =   "   Clear[F1]"
         PicturePosition =   327683
         Size            =   "3360;979"
         Picture         =   "frmReEnrol2.frx":0988
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdOk 
         Height          =   555
         Left            =   5310
         TabIndex        =   21
         Top             =   450
         Width           =   1905
         Caption         =   " 검체확인[F4]"
         PicturePosition =   327683
         Size            =   "3360;979"
         Picture         =   "frmReEnrol2.frx":211A
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
Attribute VB_Name = "frmReenrol2"
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
    
    sprConfirm.Row = 1
    sprConfirm.Row2 = sprConfirm.DataRowCnt
    sprConfirm.Col = 1
    sprConfirm.Col2 = sprConfirm.MaxCols
    sprConfirm.BlockMode = True
    sprConfirm.Action = ActionClearText
    sprConfirm.BlockMode = False

End Sub

Private Sub cmdOk_Click()
    Dim sRowID      As String
    Dim nOrderno    As String
    Dim sOrderno    As String
    
    
    If sprConfirm.DataRowCnt = 0 Then
        MsgBox "해당 접수 Data 가 하나도 없습니다!.. 접수한 Data를 먼저확인하세요"
        Exit Sub
    End If
    
    For i = 1 To sprConfirm.DataRowCnt
        sprConfirm.Row = i
        sprConfirm.Col = 1: sRowID = sprConfirm.Text
        sprConfirm.Col = 8: sOrderno = sprConfirm.Text
        GoSub Update_General_GbCH
    Next
    
    Call cmdClear_Click
    txtBarCode.SetFocus
    
    Exit Sub
    

Update_General_GbCH:
    strSql = ""
    strSql = strSql & " UPDATE TWEXAM_General"
    strSql = strSql & " SET    GbCH  = 'Y'"
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
    strSql = strSql & " UPDATE TWEXAM_Order"
    strSql = strSql & " SET    GbCH      =  'Y',"
    strSql = strSql & "        CoLLDate  =   TO_DATE('" & sToDate & "','YYYY-MM-DD'),"
    strSql = strSql & "        CoLLHH    =   " & nToHH & ","
    strSql = strSql & "        CoLLMM    =   " & nToMM & ","
    strSql = strSql & "        CoLLid    =   " & Val(GstrIdnumber)
    strSql = strSql & " WHERE  JeobsuDt  =   TO_DATE('" & sJDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    SLipno1   =  '" & Val(sSLno1) & "'"
    strSql = strSql & " AND    Orderno   =   " & Val(sOrderno) & ""
    
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    Return
    

End Sub

Private Sub mnuExit_Click()
    
    Unload Me
    
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
        cmdOk.SetFocus
        
    End If
    
    Exit Sub
    

Get_General_Data:
    strSql = ""
    strSql = strSql & " SELECT a.RowID RwID, a.*,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDt,'YYYY-MM-DD') JeobsuDt,"
    strSql = strSql & "        b.Codenm SLipName, c.Codenm SampleName, a.Orderno"
    strSql = strSql & " FROM   TWEXAM_General a,"
    strSql = strSql & "        TWEXAM_SpeCode b,"
    strSql = strSql & "        TWEXAM_Sample  c "
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
            Case "1": sprConfirm.Col = 7: sprConfirm.Text = "병동 Or ER 접수"
            Case "2": sprConfirm.Col = 7: sprConfirm.Text = "정규채혈"
        End Select
        
        txtPtno.Text = adoSet.Fields("Ptno").Value & ""
        sprConfirm.Col = 8: sprConfirm.Text = adoSet.Fields("Orderno").Value & ""
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
    strSql = strSql & "        TWBas_Room     b,"
    strSql = strSql & "        TWBas_Dept     c,"
    strSql = strSql & "        TWBas_Doctor   d "
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
