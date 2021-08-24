VERSION 5.00
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmBarCode 
   Caption         =   "BarCode Label Text"
   ClientHeight    =   3795
   ClientLeft      =   2565
   ClientTop       =   3285
   ClientWidth     =   8025
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
   ScaleHeight     =   3795
   ScaleWidth      =   8025
   Begin VB.TextBox txtDaySeq 
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      Height          =   270
      Left            =   180
      TabIndex        =   10
      Top             =   630
      Width           =   375
   End
   Begin FPSpreadADO.fpSpread ssLabel 
      Height          =   2400
      Left            =   630
      TabIndex        =   8
      Top             =   585
      Width           =   7125
      _Version        =   196608
      _ExtentX        =   12568
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
      MaxCols         =   16
      ScrollBars      =   2
      SpreadDesigner  =   "frmBarCode.frx":0000
      Appearance      =   1
   End
   Begin VB.ListBox lstSeq 
      Height          =   2040
      Left            =   180
      TabIndex        =   7
      Top             =   900
      Width           =   375
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   810
      Top             =   3105
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      Handshaking     =   1
   End
   Begin VB.TextBox txtDrname 
      BackColor       =   &H00C0E0FF&
      Height          =   330
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "txtDrname"
      Top             =   180
      Width           =   915
   End
   Begin VB.TextBox txtDrcode 
      BackColor       =   &H00C0E0FF&
      Height          =   330
      Left            =   5175
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Drcode"
      Top             =   180
      Width           =   915
   End
   Begin VB.TextBox txtRoom 
      BackColor       =   &H00C0E0FF&
      Height          =   330
      Left            =   4230
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "RoomCode"
      Top             =   180
      Width           =   915
   End
   Begin VB.TextBox txtAge 
      BackColor       =   &H00C0E0FF&
      Height          =   330
      Left            =   3735
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "Age"
      Top             =   180
      Width           =   510
   End
   Begin VB.TextBox txtSex 
      BackColor       =   &H00C0E0FF&
      Height          =   330
      Left            =   3285
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "Sex"
      Top             =   180
      Width           =   465
   End
   Begin VB.TextBox txtSname 
      BackColor       =   &H00C0E0FF&
      Height          =   330
      Left            =   1845
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "Sname"
      Top             =   180
      Width           =   1455
   End
   Begin VB.TextBox txtPtno 
      BackColor       =   &H00C0E0FF&
      Height          =   330
      Left            =   630
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "Ptno"
      Top             =   180
      Width           =   1185
   End
   Begin MSForms.CommandButton cmdPrintOk 
      Height          =   420
      Left            =   5940
      TabIndex        =   12
      Top             =   3060
      Width           =   1815
      Caption         =   "Print"
      Size            =   "3201;741"
      FontName        =   "굴림"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdClear 
      Height          =   420
      Left            =   4185
      TabIndex        =   11
      Top             =   3060
      Width           =   1770
      Caption         =   "Clear"
      Size            =   "3122;741"
      FontName        =   "굴림"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdExecute 
      Height          =   420
      Left            =   2610
      TabIndex        =   9
      Top             =   3060
      Width           =   1590
      Caption         =   "Execute"
      Size            =   "2805;741"
      FontName        =   "굴림"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmBarCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    
    txtDaySeq.Text = ""
    ssLabel.Row = 1
    ssLabel.Row2 = ssLabel.DataRowCnt + 1
    ssLabel.Col = 1
    ssLabel.Col2 = ssLabel.DataColCnt
    ssLabel.BlockMode = True
    ssLabel.Action = ActionClear
    ssLabel.BlockMode = False
    
End Sub

Private Sub cmdExec_Click()

    
    txtPtno.Text = GLabelPtno
    GoSub Get_PatientData       '환자정보 Select
    
    GoSub Get_DaySequence       'DaySeq(Twexam_General_Sub) 에서의 Group by
    
    If Trim(GLabelJDt) = "" And Trim(GLabelJT1) = "" And Trim(GLabelJT2) = "" Then
        GoSub MainProcessing
    Else
        GoSub MainProcessing_Part   'General 의 JeobsuT1,2를 Key로 하여 부분적인 Label 발행일경우,,,
    End If
    
    GoSub ReSelect_Variable
    GoSub Display_ArrayTo_Spread
    
    Exit Sub
    
    
    
Get_PatientData:
    If IsAdmission(txtPtno.Text) Then
        GoSub Get_ADMaster
    Else
        GoSub Get_HJMaster
    End If
    Return
    


Get_ADMaster:
    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TWIPD_MASTER  INDEX_IPDMST2)   "
    'strSql = strSql & "            INDEX (TW_MIS_PMPA.TWBAS_DOCTOR  INDEX_DOCTOR0) */"
    
    strSql = ""
    strSql = strSql & " SELECT a.Sname, a.Sex, a.Age, a.RoomCode, a.DrCode, b.Drname"
    strSql = strSql & " FROM   TW_MIS_PMPA.TWIPD_MASTER a,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR b "
    strSql = strSql & " WHERE  a.Ptno   =  '" & txtPtno.Text & "'"
    strSql = strSql & " AND    a.Drcode = b.Drcode(+)"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    txtSname.Text = adoSet.Fields("Sname").Value & ""
    txtSex.Text = adoSet.Fields("Sex").Value & ""
    txtAge.Text = adoSet.Fields("Age").Value & ""
    txtRoom.Text = adoSet.Fields("RoomCode").Value & ""
    txtDrcode.Text = adoSet.Fields("Drcode").Value & ""
    txtDrname.Text = adoSet.Fields("Drname").Value & ""
    Call adoSetClose(adoSet)
    Return
    

Get_HJMaster:
    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INX_PATIENT0)    "
    'strSql = strSql & "            INDEX (TW_MIS_PMPA.TWBAS_DOCTOR  INDEX_DOCTOR0) */"
    
    strSql = ""
    strSql = strSql & " SELECT a.Sname, a.Sex, a.Jumin1, a.Jumin2, a.Drcode,"
    strSql = strSql & "        b.Drname"
    strSql = strSql & " FROM   TW_MIS_PMPA.TWBAS_PATIENT a,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR  b "
    strSql = strSql & " WHERE  a.Ptno    =  '" & txtPtno.Text & "'"
    strSql = strSql & " AND    a.Drcode  = b.Drcode(+)"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    txtSname.Text = adoSet.Fields("Sname").Value & ""
    txtSex.Text = adoSet.Fields("Sex").Value & ""
    txtRoom.Text = ""
    txtAge.Text = SetAge_Check(adoSet.Fields("Jumin1").Value & "", _
                               adoSet.Fields("Jumin2").Value & "")
    txtDrcode.Text = adoSet.Fields("Drcode").Value & ""
    txtDrname.Text = adoSet.Fields("Drname").Value & ""
    Call adoSetClose(adoSet)
    Return
    


Get_DaySequence:
    strSql = ""
    strSql = strSql & " SELECT DAYSEQ"
    strSql = strSql & " FROM   TWEXAM_GENERAL_SUB"
    strSql = strSql & " WHERE  Jeobsudt = TO_DATE('" & GLabelJeobsuDt & "','yyyy-mm-dd')"
    strSql = strSql & " AND    Ptno     = '" & GLabelPtno & "'"
    strSql = strSql & " GROUP  BY dayseq"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    lstSeq.Clear
    Do Until adoSet.EOF
        lstSeq.AddItem Val(adoSet.Fields("DAYSEQ").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return

MainProcessing:
    Dim strSum      As String
    Dim nMaxSeq     As Integer
    
    If lstSeq.ListCount = 0 Then Exit Sub
    If Trim(txtDaySeq.Text) = "" Then
        lstSeq.Selected(lstSeq.ListCount - 1) = True
    End If
    nMaxSeq = Val(txtDaySeq.Text)
    
    
    'Routine Code 의 약어를 읽지 않고 ItemCode 의 BarText만으로 BartCodePrinting....
    '연속검사의 BarCode 때문에 ...
    
    Call LabelStringClear
    'strSql = ""
    'strSql = strSql & "  SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_DEPT INDEX_DEPT0)  */"
    
    strSql = ""
    strSql = strSql & " SELECT  a.Ptno, TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt, "
    strSql = strSql & "         a.SLipno1, a.SLipno2, b.BarText, b.ChwhYg, c.GeomchCD, b.GeomsaGb,b.BarGb,"
    strSql = strSql & "         c.GbEr,"
    strSql = strSql & "         d.Deptnamek"
    strSql = strSql & "  FROM   TWEXAM_General_Sub a,"
    strSql = strSql & "         TW_MIS_EXAM.TWEXAM_itemML      b,"
    strSql = strSql & "         TWEXAM_General     c,"
    strSql = strSql & "         TW_MIS_PMPA.TWBAS_DEPT         d "
    strSql = strSql & "  WHERE  a.Ptno     =  '" & GLabelPtno & "'"
    strSql = strSql & "  AND    a.JeobsuDt = TO_DATE('" & GLabelJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & "  AND    a.DaySeq   =   " & nMaxSeq
    strSql = strSql & "  AND   ( a.Routincd = a.ItemCd Or b.BarGb = '1')"
    strSql = strSql & "  AND    a.ItemCD   = b.Codeky(+)"
    strSql = strSql & "  AND    a.JeobsuDt = c.JeobsuDt(+)"
    strSql = strSql & "  AND    a.SLipno1  = c.SLipno1(+)"
    strSql = strSql & "  AND    a.SLipno2  = c.SLipno2(+)"
    strSql = strSql & "  AND    c.DeptCode = d.DeptCode(+)"
    strSql = strSql & "  GROUP BY a.Ptno, a.JeobsuDt, a.SLipno1, a.SLipno2, b.BarText, b.Chwhyg, c.GeomchCD, "
    strSql = strSql & "           b.GeomsaGb, b.BarGb, c.GbEr, d.Deptnamek"
    strSql = strSql & " UNION ALL"
    'strSql = strSql & "  SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_DEPT INDEX_DEPT0)  */"
    strSql = strSql & " SELECT  a.Ptno, TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt, "
    strSql = strSql & "         a.SLipno1, a.SLipno2, d.YakCd BarText, b.ChwhYg, c.GeomchCD, b.GeomsaGb,b.BarGb,"
    strSql = strSql & "         c.GbEr,"
    strSql = strSql & "         e.Deptnamek"
    strSql = strSql & "  FROM   TWEXAM_General_Sub a,"
    strSql = strSql & "         TW_MIS_EXAM.TWEXAM_itemML      b,"
    strSql = strSql & "         TWEXAM_General     c,"
    strSql = strSql & "         TW_MIS_EXAM.TWEXAM_Routine     d,"
    strSql = strSql & "         TW_MIS_PMPA.TWBAS_DEPT         e "
    strSql = strSql & "  WHERE  a.Ptno      =  '" & GLabelPtno & "'"
    strSql = strSql & "  AND    a.JeobsuDt  = TO_DATE('" & GLabelJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & "  AND    a.DaySeq    =   " & nMaxSeq
    strSql = strSql & "  AND    a.Routincd != a.ItemCd "
    strSql = strSql & "  AND    a.ItemCD    = b.Codeky(+)"
    strSql = strSql & "  AND    a.JeobsuDt  = c.JeobsuDt(+)"
    strSql = strSql & "  AND    a.SLipno1   = c.SLipno1(+)"
    strSql = strSql & "  AND    a.SLipno2   = c.SLipno2(+)"
    strSql = strSql & "  AND    a.RoutinCd  = d.RoutinCD"
    strSql = strSql & "  AND   (d.Series IS NULL OR d.Series != '1')"
    strSql = strSql & "  AND    c.DeptCode  = e.DeptCode(+)"
    strSql = strSql & "  GROUP BY a.Ptno, a.JeobsuDt, a.SLipno1, a.SLipno2, d.Yakcd, b.Chwhyg, c.GeomchCD, "
    strSql = strSql & "           b.GeomsaGb, b.BarGb, c.GbEr, e.Deptnamek"

    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    i = 0
    Do Until adoSet.EOF
        LabelString.Ptno(i) = adoSet.Fields("Ptno").Value & ""
        LabelString.JeobsuDt(i) = adoSet.Fields("JeobsuDt").Value & ""
        LabelString.sLipno1(i) = adoSet.Fields("SLipno1").Value & ""
        LabelString.Slipno2(i) = adoSet.Fields("SLipno2").Value & ""
        LabelString.BarText(i) = adoSet.Fields("BarText").Value & ""
        LabelString.Yg(i) = adoSet.Fields("Chwhyg").Value & ""
        LabelString.SampleCd(i) = adoSet.Fields("GeomchCD").Value & ""
        LabelString.ReporCd(i) = adoSet.Fields("GeomsaGb").Value & ""
        LabelString.Er(i) = adoSet.Fields("GbEr").Value & ""
        LabelString.DeptCode(i) = Trim(adoSet.Fields("DeptNamek").Value & "")
        
        
        LabelString.Title(i) = LabelString.Ptno(i) & _
                               LabelString.JeobsuDt(i) & _
                               LabelString.sLipno1(i) & _
                               LabelString.Slipno2(i) & _
                               LabelString.Yg(i) & _
                               LabelString.SampleCd(i) & _
                               LabelString.ReporCd(i) & _
                               LabelString.Er(i)
        
        If adoSet.Fields("BarGB").Value & "" = "1" Then            'BarCode Label 을 따로 관리하는 항목은 ....
            LabelString.Title(i) = LabelString.Title(i) & LabelString.BarText(i)
        End If
                               
        adoSet.MoveNext: i = i + 1
    Loop
    Call adoSetClose(adoSet)
    
    Return
    
    
MainProcessing_Part:
    
    'JeobsuDt, JeobsuT1, JeobsuT2 의 조건으로 화면을 Load 하였을경우, 선택적인 BarCode Label
    'Routine Code 의 약어를 읽지 않고 ItemCode 의 BarText만으로 BartCodePrinting....
    '연속검사의 BarCode 때문에 ...
    
    Call LabelStringClear
    
    'strSql = ""
    'strSql = strSql & "  SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_DEPT INDEX_DEPT0)  */"
    
    strSql = ""
    strSql = strSql & " SELECT  a.Ptno, TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt, "
    strSql = strSql & "         a.SLipno1, a.SLipno2, b.BarText, b.ChwhYg, c.GeomchCD, b.GeomsaGb, "
    strSql = strSql & "         b.BarGB, c.GbEr,"
    strSql = strSql & "         d.Deptnamek"
    strSql = strSql & "  FROM   TWEXAM_General_Sub a,"
    strSql = strSql & "         TW_MIS_EXAM.TWEXAM_itemML      b,"
    strSql = strSql & "         TWEXAM_General     c,"
    strSql = strSql & "         TW_MIS_PMPA.TWBAS_DEPT         d "
    strSql = strSql & "  WHERE  a.Ptno     =  '" & GLabelPtno & "'"
    strSql = strSql & "  AND    a.JeobsuDt = TO_DATE('" & GLabelJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & "  AND    c.JeobsuDt = TO_DATE('" & GLabelJDt & "','YYYY-MM-DD')"
    strSql = strSql & "  AND    c.JeobsuT1 = '" & GLabelJT1 & "'"
    strSql = strSql & "  AND    c.JeobsuT2 = '" & GLabelJT2 & "'"
    strSql = strSql & "  AND    a.ItemCd   = a.Itemcd"
    strSql = strSql & "  AND    a.ItemCD   = b.Codeky(+)"
    strSql = strSql & "  AND    a.JeobsuDt = c.JeobsuDt(+)"
    strSql = strSql & "  AND    a.SLipno1  = c.SLipno1(+)"
    strSql = strSql & "  AND    a.SLipno2  = c.SLipno2(+)"
    strSql = strSql & "  AND    c.DeptCode = d.DeptCode(+)"
    strSql = strSql & "  GROUP BY a.Ptno, a.JeobsuDt, a.SLipno1, a.SLipno2, b.BarText, b.Chwhyg, "
    strSql = strSql & "           b.BarGb, c.GeomchCD, b.GeomsaGb, c.GbEr, d.Deptnamek"

    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    i = 0
    Do Until adoSet.EOF
        LabelString.Ptno(i) = adoSet.Fields("Ptno").Value & ""
        LabelString.JeobsuDt(i) = adoSet.Fields("JeobsuDt").Value & ""
        LabelString.sLipno1(i) = adoSet.Fields("SLipno1").Value & ""
        LabelString.Slipno2(i) = adoSet.Fields("SLipno2").Value & ""
        LabelString.BarText(i) = adoSet.Fields("BarText").Value & ""
        LabelString.Yg(i) = adoSet.Fields("Chwhyg").Value & ""
        LabelString.SampleCd(i) = adoSet.Fields("GeomchCD").Value & ""
        LabelString.ReporCd(i) = adoSet.Fields("GeomsaGb").Value & ""
        LabelString.Er(i) = adoSet.Fields("GbEr").Value & ""
        LabelString.DeptCode(i) = Trim(adoSet.Fields("DeptNamek").Value & "")
        
        
        LabelString.Title(i) = LabelString.Ptno(i) & _
                               LabelString.JeobsuDt(i) & _
                               LabelString.sLipno1(i) & _
                               LabelString.Slipno2(i) & _
                               LabelString.Yg(i) & _
                               LabelString.SampleCd(i) & _
                               LabelString.ReporCd(i) & _
                               LabelString.Er(i)
        
        If adoSet.Fields("BarGB").Value & "" = "1" Then            'BarCode Label 을 따로 관리하는 항목은 ....
            LabelString.Title(i) = LabelString.Title(i) & LabelString.BarText(i)
        End If
        
        adoSet.MoveNext: i = i + 1
    Loop
    Call adoSetClose(adoSet)
    
    Return
    
ReSelect_Variable:
    Dim nStart       As String
    
    Call LabelString1Clear
    
    For i = 0 To 50
        If isArrayText(LabelString1.Title, LabelString.Title(i)) Then
            If Trim(LabelString.BarText(i)) <> "" Then
                LabelString1.BarText(GVarPoint) = LabelString1.BarText(GVarPoint) & "," & LabelString.BarText(i)
            End If
        Else
            nStart = isArrayMaxReturn(LabelString1.Title)
            LabelString1.Title(nStart) = LabelString.Title(i)
            LabelString1.Ptno(nStart) = LabelString.Ptno(i)
            LabelString1.JeobsuDt(nStart) = LabelString.JeobsuDt(i)
            LabelString1.sLipno1(nStart) = LabelString.sLipno1(i)
            LabelString1.Slipno2(nStart) = LabelString.Slipno2(i)
            LabelString1.BarText(nStart) = LabelString.BarText(i)
            LabelString1.Yg(nStart) = LabelString.Yg(i)
            LabelString1.SampleCd(nStart) = LabelString.SampleCd(i)
            LabelString1.ReporCd(nStart) = LabelString.ReporCd(i)
            LabelString1.Er(nStart) = LabelString.Er(i)
            LabelString1.DeptCode(nStart) = LabelString.DeptCode(i)
        End If
    Next
    Return


Display_ArrayTo_Spread:
    Call Spread_Set_Clear(ssLabel)
    
    For i = 0 To 100
        ssLabel.Row = i + 1
        If LabelString1.Title(i) <> "" Then
            ssLabel.Col = 1:  ssLabel.Value = True
            ssLabel.Col = 2:  ssLabel.Text = LabelString1.Ptno(i)
            ssLabel.Col = 3:  ssLabel.Text = txtSname.Text
            ssLabel.Col = 4:  ssLabel.Text = txtRoom.Text
                              
                                 
                                
            ssLabel.Col = 5:  ssLabel.Text = LabelString1.JeobsuDt(i)
            ssLabel.Col = 6:  ssLabel.Text = LabelString1.sLipno1(i)
            ssLabel.Col = 7:  ssLabel.Text = Format(LabelString1.Slipno2(i), "00000")
            ssLabel.Col = 8:  ssLabel.Text = LabelString1.BarText(i)
            ssLabel.Col = 9:  ssLabel.TypeComboBoxCurSel = 1
            ssLabel.Col = 10: ssLabel.Text = LabelString1.SampleCd(i)
                              GoSub Get_SampleData
            ssLabel.Col = 12: ssLabel.Text = LabelString1.Yg(i)
                              'GoSub Get_YgData
            ssLabel.Col = 14: ssLabel.Text = LabelString1.Er(i)
            ssLabel.Col = 15: ssLabel.Text = LabelString1.ReporCd(i)
            ssLabel.Col = 16: ssLabel.Text = LabelString1.DeptCode(i)
            'GoSub Get_Emergency_Check
            
        End If
    Next
    
    Return
    
    
Get_SampleData:
    strSql = ""
    strSql = strSql & " SELECT * "
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Sample"
    strSql = strSql & " WHERE  Code = '" & LabelString1.SampleCd(i) & "'"
    If False = adoSetOpen(strSql, adoSet) Then Return
    ssLabel.Col = 11: ssLabel.Text = Trim(adoSet.Fields("Codenm").Value & "")
    Call adoSetClose(adoSet)
    Return

Get_YgData:
    strSql = ""
    strSql = strSql & " SELECT CODENM, Yageo"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Specode"
    strSql = strSql & " WHERE  CODEGU = '88'"
    strSql = strSql & " AND    CODEKY = '" & LabelString1.Yg(i) & "'"
    If False = adoSetOpen(strSql, adoSet) Then Return
    ssLabel.Col = 13: ssLabel.Text = Trim(adoSet.Fields("Yageo").Value & "")
    Call adoSetClose(adoSet)
    Return

Get_Emergency_Check:
    strSql = ""
    strSql = strSql & " SELECT GBER"
    strSql = strSql & " FROM   TWEXAM_General"
    strSql = strSql & " WHERE  JeobsuDt =   TO_DATE('" & LabelString1.JeobsuDt(i) & "','YYYY-MM-DD')"
    strSql = strSql & " AND    SLipno1  =   " & Val(LabelString1.sLipno1(i))
    strSql = strSql & " AND    SLipno2  =   " & Val(LabelString1.Slipno2(i))
    If False = adoSetOpen(strSql, adoSet) Then Return
    If Trim(adoSet.Fields("GbEr").Value & "") <> "" Then
        ssLabel.Col = 14: ssLabel.Text = adoSet.Fields("GbEr").Value & ""
    End If
    Call adoSetClose(adoSet)
    Return

End Sub

Public Sub cmdExecute_Click()

    
    txtPtno.Text = GLabelPtno
    GoSub Get_PatientData       '환자정보 Select
    
    GoSub Get_DaySequence       'DaySeq(Twexam_General_Sub) 에서의 Group by
    
    If Trim(GLabelJDt) = "" And Trim(GLabelJT1) = "" And Trim(GLabelJT2) = "" Then
        GoSub MainProcessing
    Else
        GoSub MainProcessing_Part   'General 의 JeobsuT1,2를 Key로 하여 부분적인 Label 발행일경우,,,
    End If
    
    GoSub ReSelect_Variable
    GoSub Display_ArrayTo_Spread
    
    Exit Sub
    
    
    
Get_PatientData:
    If IsAdmission(txtPtno.Text) Then
        GoSub Get_ADMaster
    Else
        GoSub Get_HJMaster
    End If
    Return
    


Get_ADMaster:
    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TWIPD_MASTER  INDEX_IPDMST2)   "
    'strSql = strSql & "            INDEX (TW_MIS_PMPA.TWBAS_DOCTOR  INDEX_DOCTOR0) */"
    
    '입원환자일 경우 입원 MASTER FILE에서 환자성명/성별/나이/병실/의사코드/의사성명의 자료를 가지고 온다.
    strSql = ""
    strSql = strSql & " SELECT a.Sname, a.Sex, a.Age, a.RoomCode, a.DrCode, b.Drname"
    strSql = strSql & " FROM   TW_MIS_PMPA.TWIPD_MASTER a,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR b "
    strSql = strSql & " WHERE  a.Ptno   =  '" & txtPtno.Text & "'"
    strSql = strSql & " AND    a.Drcode = b.Drcode(+)"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    txtSname.Text = adoSet.Fields("Sname").Value & ""
    txtSex.Text = adoSet.Fields("Sex").Value & ""
    txtAge.Text = adoSet.Fields("Age").Value & ""
    txtRoom.Text = adoSet.Fields("RoomCode").Value & ""
    txtDrcode.Text = adoSet.Fields("Drcode").Value & ""
    txtDrname.Text = adoSet.Fields("Drname").Value & ""
    Call adoSetClose(adoSet)
    Return
    

Get_HJMaster:
    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INX_PATIENT0)    "
    'strSql = strSql & "            INDEX (TW_MIS_PMPA.TWBAS_DOCTOR  INDEX_DOCTOR0) */"
    
    '외래환자일 경우 환자기초자료에서 이름/성별/주민번호/의사코드/의사성명을 가지고 온다.
    strSql = ""
    strSql = strSql & " SELECT a.Sname, a.Sex, a.Jumin1, a.Jumin2, a.Drcode,"
    strSql = strSql & "        b.Drname"
    strSql = strSql & " FROM   TW_MIS_PMPA.TWBAS_PATIENT a,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR  b "
    strSql = strSql & " WHERE  a.Ptno    =  '" & txtPtno.Text & "'"
    strSql = strSql & " AND    a.Drcode  = b.Drcode(+)"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    txtSname.Text = adoSet.Fields("Sname").Value & ""
    txtSex.Text = adoSet.Fields("Sex").Value & ""
    txtRoom.Text = ""
    txtAge.Text = SetAge_Check(adoSet.Fields("Jumin1").Value & "", _
                               adoSet.Fields("Jumin2").Value & "")
    txtDrcode.Text = adoSet.Fields("Drcode").Value & ""
    txtDrname.Text = adoSet.Fields("Drname").Value & ""
    Call adoSetClose(adoSet)
    Return
    


Get_DaySequence:
    lstSeq.Clear
    strSql = ""
    strSql = strSql & " SELECT DAYSEQ"
    strSql = strSql & " FROM   TWEXAM_GENERAL_SUB"
    strSql = strSql & " WHERE  Jeobsudt = TO_DATE('" & GLabelJeobsuDt & "','yyyy-mm-dd')"
    strSql = strSql & " AND    Ptno     = '" & GLabelPtno & "'"
    strSql = strSql & " GROUP  BY dayseq"
    
    
    If False = adoSetOpen(strSql, adoSet) Then
        lstSeq.AddItem "0"
        Return
    End If
    
    Do Until adoSet.EOF
        lstSeq.AddItem Val(adoSet.Fields("DAYSEQ").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return

MainProcessing:
    Dim strSum      As String
    Dim nMaxSeq     As Integer
    
    If lstSeq.ListCount = 0 Then Exit Sub
    If lstSeq.ListCount > 0 Then
        If Trim(txtDaySeq.Text) = "" Then
            lstSeq.Selected(lstSeq.ListCount - 1) = True
        End If
    End If
    nMaxSeq = Val(txtDaySeq.Text)
    
    
    'Routine Code 의 약어를 읽지 않고 ItemCode 의 BarText만으로 BartCodePrinting....
    '연속검사의 BarCode 때문에 ...
    
    Call LabelStringClear
    
    strSql = ""
    strSql = strSql & "  SELECT "
    strSql = strSql & "         a.Ptno, TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt, "
    strSql = strSql & "         a.SLipno1, a.SLipno2, b.BarText, b.ChwhYg, e.GeomchCD, b.GeomsaGb,b.BarGb,"
    strSql = strSql & "         e.GbEr,b.ChUnit,"
    strSql = strSql & "         d.Deptnamek"
    strSql = strSql & "  FROM   TWEXAM_General_Sub             a,"
    strSql = strSql & "         TW_MIS_EXAM.TWEXAM_itemML      b,"
    strSql = strSql & "         TWEXAM_General                 c,"
    strSql = strSql & "         TW_MIS_PMPA.TWBAS_DEPT         d,"
    strSql = strSql & "         TW_MIS_EXAM.TWEXAM_Order       e "
    strSql = strSql & "  WHERE  a.Ptno     =  '" & GLabelPtno & "'"
    strSql = strSql & "  AND    a.JeobsuDt = TO_DATE('" & GLabelJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & "  AND    a.DaySeq   =   " & nMaxSeq
    strSql = strSql & "  AND   ( a.Routincd = a.ItemCd Or b.BarGb = '1')"
    strSql = strSql & "  AND    a.ItemCD   = b.Codeky"
    strSql = strSql & "  AND    a.JeobsuDt = c.JeobsuDt(+)"
    strSql = strSql & "  AND    a.SLipno1  = c.SLipno1(+)"
    strSql = strSql & "  AND    a.SLipno2  = c.SLipno2(+)"
    strSql = strSql & "  AND    a.JeobsuDt = e.CollDate(+)"
    strSql = strSql & "  AND    a.Orderno  = e.Orderno(+)"
    strSql = strSql & "  AND    a.SLipno1  = e.SLipno1"
    strSql = strSql & "  AND    e.DeptCode = d.DeptCode(+)"
    strSql = strSql & "  GROUP BY a.Ptno, a.JeobsuDt, a.SLipno1, a.SLipno2, b.BarText, b.Chwhyg, e.GeomchCD, "
    strSql = strSql & "           b.GeomsaGb, b.BarGb, e.GbEr, b.ChUnit, d.Deptnamek"
    strSql = strSql & " UNION ALL"
    strSql = strSql & "  SELECT "
    strSql = strSql & "         a.Ptno, TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt, "
    strSql = strSql & "         a.SLipno1, a.SLipno2, d.YakCd BarText, b.ChwhYg, f.GeomchCD, b.GeomsaGb,b.BarGb,"
    strSql = strSql & "         f.GbEr, b.ChUnit,"
    strSql = strSql & "         e.Deptnamek"
    strSql = strSql & "  FROM   TWEXAM_General_Sub             a,"
    strSql = strSql & "         TW_MIS_EXAM.TWEXAM_itemML      b,"
    strSql = strSql & "         TW_MIS_EXAM.TWEXAM_Routine     d,"
    strSql = strSql & "         TW_MIS_PMPA.TWBAS_DEPT         e,"
    strSql = strSql & "         TW_MIS_EXAM.TWEXAM_Order       f "
    strSql = strSql & "  WHERE  a.Ptno      =  '" & GLabelPtno & "'"
    strSql = strSql & "  AND    a.JeobsuDt  = TO_DATE('" & GLabelJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & "  AND    a.DaySeq    =   " & nMaxSeq
    strSql = strSql & "  AND    a.Routincd != a.ItemCd "
    strSql = strSql & "  AND    a.ItemCD    = b.Codeky"
    strSql = strSql & "  AND   (b.BarGb IS NULL OR b.BarGb != '1')"
    strSql = strSql & "  AND    a.RoutinCd  = d.RoutinCD"
    strSql = strSql & "  AND   (d.Series IS NULL OR d.Series != '1')"
    strSql = strSql & "  AND    a.JeobsuDt  = f.CollDate(+)"
    strSql = strSql & "  AND    a.Orderno   = f.Orderno(+)"
    strSql = strSql & "  AND    a.SLipno1   = f.SLipno1"
    strSql = strSql & "  AND    f.DeptCode  = e.DeptCode(+)"
    strSql = strSql & "  GROUP BY a.Ptno, a.JeobsuDt, a.SLipno1, a.SLipno2, d.Yakcd, b.Chwhyg, f.GeomchCD, "
    strSql = strSql & "           b.GeomsaGb, b.BarGb, f.GbEr, b.ChUnit,e.Deptnamek"

    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    i = 0
    Do Until adoSet.EOF
        LabelString.Ptno(i) = adoSet.Fields("Ptno").Value & ""
        LabelString.JeobsuDt(i) = adoSet.Fields("JeobsuDt").Value & ""
        LabelString.sLipno1(i) = adoSet.Fields("SLipno1").Value & ""
        LabelString.Slipno2(i) = adoSet.Fields("SLipno2").Value & ""
        LabelString.BarText(i) = adoSet.Fields("BarText").Value & ""
        LabelString.Yg(i) = adoSet.Fields("Chwhyg").Value & ""
        LabelString.SampleCd(i) = adoSet.Fields("GeomchCD").Value & ""
        LabelString.ReporCd(i) = adoSet.Fields("GeomsaGb").Value & ""
        LabelString.Er(i) = adoSet.Fields("GbEr").Value & ""
        LabelString.DeptCode(i) = Trim(adoSet.Fields("DeptNamek").Value & "")
        LabelString.ChUnit(i) = Trim(adoSet.Fields("ChUnit").Value & "")
        
        LabelString.Title(i) = LabelString.Ptno(i) & _
                               LabelString.JeobsuDt(i) & _
                               LabelString.sLipno1(i) & _
                               LabelString.Slipno2(i) & _
                               LabelString.Yg(i) & _
                               LabelString.SampleCd(i) & _
                               LabelString.ReporCd(i) & _
                               LabelString.Er(i)
        
        If adoSet.Fields("BarGB").Value & "" = "1" Then            'BarCode Label 을 따로 관리하는 항목은 ....
            LabelString.Title(i) = LabelString.Title(i) & LabelString.BarText(i)
        End If
                               
        adoSet.MoveNext: i = i + 1
    Loop
    Call adoSetClose(adoSet)
    
    Return
    
    
MainProcessing_Part:
    
    'JeobsuDt, JeobsuT1, JeobsuT2 의 조건으로 화면을 Load 하였을경우, 선택적인 BarCode Label
    'Routine Code 의 약어를 읽지 않고 ItemCode 의 BarText만으로 BartCodePrinting....
    '연속검사의 BarCode 때문에 ...
    
    Call LabelStringClear
    
    
    strSql = ""
    strSql = strSql & "  SELECT "
    strSql = strSql & "         a.Ptno, TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt, "
    strSql = strSql & "         a.SLipno1, a.SLipno2, b.BarText, b.ChwhYg, e.GeomchCD, b.GeomsaGb,b.BarGb,"
    strSql = strSql & "         e.GbEr, b.ChUnit,"
    strSql = strSql & "         d.Deptnamek"
    strSql = strSql & "  FROM   TWEXAM_General_Sub a,"
    strSql = strSql & "         TW_MIS_EXAM.TWEXAM_itemML      b,"
    strSql = strSql & "         TWEXAM_General     c,"
    strSql = strSql & "         TW_MIS_PMPA.TWBAS_DEPT         d,"
    strSql = strSql & "         TW_MIS_EXAM.TWEXAM_Order       e "
    strSql = strSql & "  WHERE  a.Ptno     =  '" & GLabelPtno & "'"
    strSql = strSql & "  AND    a.JeobsuDt = TO_DATE('" & GLabelJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & "  AND    c.JeobsuDt = TO_DATE('" & GLabelJDt & "','YYYY-MM-DD')"
    strSql = strSql & "  AND    c.JeobsuT1 = '" & GLabelJT1 & "'"
    strSql = strSql & "  AND    c.JeobsuT2 = '" & GLabelJT2 & "'"
    strSql = strSql & "  AND   ( a.Routincd = a.ItemCd Or b.BarGb = '1')"
    strSql = strSql & "  AND    a.ItemCD   = b.Codeky"
    strSql = strSql & "  AND    a.JeobsuDt = c.JeobsuDt(+)"
    strSql = strSql & "  AND    a.SLipno1  = c.SLipno1(+)"
    strSql = strSql & "  AND    a.SLipno2  = c.SLipno2(+)"
    strSql = strSql & "  AND    c.DeptCode = d.DeptCode(+)"
    strSql = strSql & "  AND    a.JeobsuDt = e.CollDate(+)"
    strSql = strSql & "  AND    a.Orderno  = e.Orderno(+)"
    strSql = strSql & "  GROUP BY a.Ptno, a.JeobsuDt, a.SLipno1, a.SLipno2, b.BarText, b.Chwhyg, e.GeomchCD, "
    strSql = strSql & "           b.GeomsaGb, b.BarGb, e.GbEr, b.ChUnit, d.Deptnamek"
    strSql = strSql & " UNION ALL"
    strSql = strSql & "  SELECT "
    strSql = strSql & "         a.Ptno, TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt, "
    strSql = strSql & "         a.SLipno1, a.SLipno2, d.YakCd BarText, b.ChwhYg, f.GeomchCD, b.GeomsaGb,b.BarGb,"
    strSql = strSql & "         f.GbEr, b.ChUnit,"
    strSql = strSql & "         e.Deptnamek"
    strSql = strSql & "  FROM   TWEXAM_General_Sub a,"
    strSql = strSql & "         TW_MIS_EXAM.TWEXAM_itemML      b,"
    strSql = strSql & "         TWEXAM_General     c,"
    strSql = strSql & "         TW_MIS_EXAM.TWEXAM_Routine     d,"
    strSql = strSql & "         TW_MIS_PMPA.TWBAS_DEPT         e,"
    strSql = strSql & "         TW_MIS_EXAM.TWEXAM_Order       f "
    strSql = strSql & "  WHERE  a.Ptno     =  '" & GLabelPtno & "'"
    strSql = strSql & "  AND    a.JeobsuDt = TO_DATE('" & GLabelJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & "  AND    c.JeobsuDt = TO_DATE('" & GLabelJDt & "','YYYY-MM-DD')"
    strSql = strSql & "  AND    c.JeobsuT1 = '" & GLabelJT1 & "'"
    strSql = strSql & "  AND    c.JeobsuT2 = '" & GLabelJT2 & "'"
    strSql = strSql & "  AND    a.Routincd != a.ItemCd "
    strSql = strSql & "  AND    a.ItemCD    = b.Codeky"
    strSql = strSql & "  AND   (b.BarGb IS NULL OR b.BarGb != '1')    "
    strSql = strSql & "  AND    a.JeobsuDt  = c.JeobsuDt(+)"
    strSql = strSql & "  AND    a.SLipno1   = c.SLipno1(+)"
    strSql = strSql & "  AND    a.SLipno2   = c.SLipno2(+)"
    strSql = strSql & "  AND    a.RoutinCd  = d.RoutinCD"
    strSql = strSql & "  AND   (d.Series IS NULL OR d.Series != '1')"
    strSql = strSql & "  AND    a.JeobsuDt  = f.CollDate(+)"
    strSql = strSql & "  AND    a.Orderno   = f.Orderno(+)"
    strSql = strSql & "  AND    c.DeptCode  = e.DeptCode(+)"
    strSql = strSql & "  GROUP BY a.Ptno, a.JeobsuDt, a.SLipno1, a.SLipno2, d.Yakcd, b.Chwhyg, f.GeomchCD, "
    strSql = strSql & "           b.GeomsaGb, b.BarGb, f.GbEr, b.ChUnit, e.Deptnamek"
    
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    i = 0
    Do Until adoSet.EOF
        LabelString.Ptno(i) = adoSet.Fields("Ptno").Value & ""
        LabelString.JeobsuDt(i) = adoSet.Fields("JeobsuDt").Value & ""
        LabelString.sLipno1(i) = adoSet.Fields("SLipno1").Value & ""
        LabelString.Slipno2(i) = adoSet.Fields("SLipno2").Value & ""
        LabelString.BarText(i) = adoSet.Fields("BarText").Value & ""
        LabelString.Yg(i) = adoSet.Fields("Chwhyg").Value & ""
        LabelString.SampleCd(i) = adoSet.Fields("GeomchCD").Value & ""
        LabelString.ReporCd(i) = adoSet.Fields("GeomsaGb").Value & ""
        LabelString.Er(i) = adoSet.Fields("GbEr").Value & ""
        LabelString.DeptCode(i) = Trim(adoSet.Fields("DeptNamek").Value & "")
        LabelString.ChUnit(i) = Trim(adoSet.Fields("ChUnit").Value & "")      'BarCode Print 매수로 쓰임
        
        LabelString.Title(i) = LabelString.Ptno(i) & _
                               LabelString.JeobsuDt(i) & _
                               LabelString.sLipno1(i) & _
                               LabelString.Slipno2(i) & _
                               LabelString.Yg(i) & _
                               LabelString.SampleCd(i) & _
                               LabelString.ReporCd(i) & _
                               LabelString.Er(i)
        
        If adoSet.Fields("BarGB").Value & "" = "1" Then            'BarCode Label 을 따로 관리하는 항목은 ....
            LabelString.Title(i) = LabelString.Title(i) & LabelString.BarText(i)
        End If
        
        adoSet.MoveNext: i = i + 1
    Loop
    Call adoSetClose(adoSet)
    
    Return
    
ReSelect_Variable:
    Dim nStart       As String
    
    Call LabelString1Clear
    
    For i = 0 To 100
        If isArrayText(LabelString1.Title, LabelString.Title(i)) Then
            If Trim(LabelString.BarText(i)) <> "" Then
                LabelString1.BarText(GVarPoint) = LabelString1.BarText(GVarPoint) & "," & LabelString.BarText(i)
            End If
        Else
            nStart = isArrayMaxReturn(LabelString1.Title)
            LabelString1.Title(nStart) = LabelString.Title(i)
            LabelString1.Ptno(nStart) = LabelString.Ptno(i)
            LabelString1.JeobsuDt(nStart) = LabelString.JeobsuDt(i)
            LabelString1.sLipno1(nStart) = LabelString.sLipno1(i)
            LabelString1.Slipno2(nStart) = LabelString.Slipno2(i)
            LabelString1.BarText(nStart) = LabelString.BarText(i)
            LabelString1.Yg(nStart) = LabelString.Yg(i)
            LabelString1.SampleCd(nStart) = LabelString.SampleCd(i)
            LabelString1.ReporCd(nStart) = LabelString.ReporCd(i)
            LabelString1.Er(nStart) = LabelString.Er(i)
            LabelString1.DeptCode(nStart) = LabelString.DeptCode(i)
            LabelString1.ChUnit(nStart) = LabelString.ChUnit(i)
        End If
    Next
    Return


Display_ArrayTo_Spread:
    Call Spread_Set_Clear(ssLabel)
    
    For i = 0 To 100
        ssLabel.Row = i + 1
        If LabelString1.Title(i) <> "" Then
            ssLabel.Col = 1:  ssLabel.Value = True
            ssLabel.Col = 2:  ssLabel.Text = LabelString1.Ptno(i)
            ssLabel.Col = 3:  ssLabel.Text = txtSname.Text
            ssLabel.Col = 4:  ssLabel.Text = txtRoom.Text
                                
            ssLabel.Col = 5:  ssLabel.Text = LabelString1.JeobsuDt(i)
            ssLabel.Col = 6:  ssLabel.Text = LabelString1.sLipno1(i)
            ssLabel.Col = 7:  ssLabel.Text = Format(LabelString1.Slipno2(i), "00000")
            ssLabel.Col = 8:  ssLabel.Text = LabelString1.BarText(i)
            ssLabel.Col = 9:  ssLabel.TypeComboBoxCurSel = Val(LabelString1.ChUnit(i))
            
            ssLabel.Col = 10: ssLabel.Text = LabelString1.SampleCd(i)
                              GoSub Get_SampleData
            ssLabel.Col = 12: ssLabel.Text = LabelString1.Yg(i)
                              'GoSub Get_YgData
            ssLabel.Col = 14: ssLabel.Text = LabelString1.Er(i)
            ssLabel.Col = 15: ssLabel.Text = LabelString1.ReporCd(i)
            ssLabel.Col = 16: ssLabel.Text = LabelString1.DeptCode(i)
            'GoSub Get_Emergency_Check
            
        End If
    Next
    
    Return
    
    
Get_SampleData:
    strSql = ""
    strSql = strSql & " SELECT * "
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Sample"
    strSql = strSql & " WHERE  Code = '" & LabelString1.SampleCd(i) & "'"
    If False = adoSetOpen(strSql, adoSet) Then Return
    ssLabel.Col = 11: ssLabel.Text = Trim(adoSet.Fields("Codenm").Value & "")
    Call adoSetClose(adoSet)
    Return

Get_YgData:
    strSql = ""
    strSql = strSql & " SELECT CODENM, Yageo"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Specode"
    strSql = strSql & " WHERE  CODEGU = '88'"                                        '검체 용기
    strSql = strSql & " AND    CODEKY = '" & LabelString1.Yg(i) & "'"
    If False = adoSetOpen(strSql, adoSet) Then Return
    ssLabel.Col = 13: ssLabel.Text = Trim(adoSet.Fields("Yageo").Value & "")
    Call adoSetClose(adoSet)
    Return

Get_Emergency_Check:
    strSql = ""
    strSql = strSql & " SELECT GBER"                                                '응급검사항목 검색
    strSql = strSql & " FROM   TWEXAM_General"
    strSql = strSql & " WHERE  JeobsuDt =   TO_DATE('" & LabelString1.JeobsuDt(i) & "','YYYY-MM-DD')"
    strSql = strSql & " AND    SLipno1  =   " & Val(LabelString1.sLipno1(i))
    strSql = strSql & " AND    SLipno2  =   " & Val(LabelString1.Slipno2(i))
    If False = adoSetOpen(strSql, adoSet) Then Return
    If Trim(adoSet.Fields("GbEr").Value & "") <> "" Then
        ssLabel.Col = 14: ssLabel.Text = adoSet.Fields("GbEr").Value & ""
    End If
    Call adoSetClose(adoSet)
    Return
    
End Sub

Private Sub cmdPrintOk_Click()
    
    Dim sBarCodeText(8) As String
    
    Dim sBarSLno1       As String
    Dim sBarSLno2       As String
    Dim sBarJdate       As String
    Dim sBarText        As String
    Dim nLoop           As Integer
    Dim sSLName         As String
    Dim sBarRoom        As String
    Dim sEr             As String
    Dim sEx             As String
    Dim sSample         As String
    Dim sDeptCode       As String
    Dim sSLipText       As String
    Dim sCollDate       As String
    Dim iVar            As Integer
    
    If ssLabel.DataRowCnt = 0 Then
        MsgBox "Barcode Printing 할 Data 가 하나도 없습니다!.."
        Exit Sub
    End If
    
    For i = 1 To ssLabel.DataRowCnt
        
        ssLabel.Row = i
        ssLabel.Col = 1
        If ssLabel.Value = True Then
            GoSub Set_Array_Clear
            ssLabel.Col = 16: sDeptCode = Trim(ssLabel.Text)
            ssLabel.Col = 15: sEx = Trim(ssLabel.Text)
            ssLabel.Col = 14: sEr = Trim(ssLabel.Text)
            ssLabel.Col = 11: sSample = Trim(ssLabel.Text)
            ssLabel.Col = 6:  sBarSLno1 = ssLabel.Text
            ssLabel.Col = 7:  sBarSLno2 = Format(ssLabel.Text, "00000")
            ssLabel.Col = 8:  sBarText = ssLabel.Text
            ssLabel.Col = 5:  sBarJdate = ssLabel.Text
            ssLabel.Col = 9:  nLoop = Val(ssLabel.Text)  'Print 장수
            ssLabel.Col = 4:  sBarRoom = Trim(ssLabel.Text)   '병실Code
            
            'GoSub GET_SLipname
            sBarJdate = Replace(sBarJdate, "-", "", 1, , vbTextCompare)
            sSLipText = convSLipYageo(sBarSLno1)
            
            sBarCodeText(0) = sSLipText
            sBarCodeText(1) = sSample
            
            If Trim(sEr) <> "" Then
                sBarCodeText(2) = "응급": End If
                
            If Trim(sEx) = "W" Then
                If Trim(sBarCodeText(2)) = "" Then
                    sBarCodeText(2) = "(외)"
                Else
                    sBarCodeText(2) = sBarCodeText(2) & "/" & "(외)"
                End If
            End If
                sBarCodeText(3) = sBarJdate & "-" & sSLipText & "  " & sBarSLno2
            sBarCodeText(4) = sBarJdate & sBarSLno1 & sBarSLno2
            sBarCodeText(5) = txtPtno.Text
            sBarCodeText(6) = txtSname.Text
            'sBarCodeText(5) = txtPtno.Text & "," & txtSname.Text & "," & txtSex.Text & "/" & txtAge.Text

            If Trim(sBarRoom) = "" Then
                sBarCodeText(6) = sBarCodeText(6) & "," & sDeptCode
            Else
                sBarCodeText(6) = sBarCodeText(6) & "," & sBarRoom
            End If
                
            sBarCodeText(7) = sBarText
            Call BarCodePrint(sBarCodeText, nLoop, Me)
            
'
'            'sBarCodeText(2) = "응급/(외)"
'            'sCOLLDate = Replace(GET_COLLDate(sBarJdate, Val(sBarSLno1), Val(sBarSLno2)), "-", "")
'
'            sBarCodeText(3) = sBarJdate & "-" & sSLipText & "  " & sBarSLno2
'            'sBarCodeText(4) = sBarJdate & "-" & sBarSLno1 & sBarSLno2 '16자리
'            sBarCodeText(4) = sBarJdate & sBarSLno1 & sBarSLno2 '15자리
'
'            sBarCodeText(5) = txtPtno.Text & "," & txtSname.Text
'            'sBarCodeText(5) = txtPtno.Text & "," & txtSname.Text & "," & txtSex.Text & "/" & txtAge.Text
'
'            If Trim(sBarRoom) = "" Then
'                sBarCodeText(5) = sBarCodeText(5) & "," & sDeptCode
'            Else
'                sBarCodeText(5) = sBarCodeText(5) & "," & sBarRoom
'            End If
'
'            sBarCodeText(6) = sBarText
'            Call Bar7421_Printing_Sub(sBarCodeText, nLoop, MSComm1)
        End If
    Next
    
    If GLabelLoadCheck = "" Then
        Me.Hide
        Unload Me
    End If
    Exit Sub
    

    
    
Set_Array_Clear:
    
    
    For iVar = 0 To 7
        sBarCodeText(iVar) = ""
    Next
    
    sEx = ""
    sEr = ""
    sBarSLno1 = ""
    sBarSLno2 = ""
    sBarText = ""
    sBarJdate = ""
    nLoop = 0
    sBarRoom = ""
    
    Return
    
    
GET_SLipname:
    strSql = " SELECT Yageo FROM TW_MIS_EXAM.TWEXAM_Specode WHERE CODEGU = '12' AND Codeky = '" & sBarSLno1 & "'"
    If False = adoSetOpen(strSql, adoSet) Then
        sSLName = ""
        Return
    End If
    sSLName = "[" & Trim(adoSet.Fields("Yageo").Value & "") & "]"
    Call adoSetClose(adoSet)
    Return
    
End Sub

Private Sub Form_Activate()
    
    DoEvents: Call cmdExecute_Click
    If GLabelLoadCheck = "" Then
        DoEvents: Call cmdPrintOk_Click
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    GLabelLoadCheck = ""
    GLabelJDt = ""
    GLabelJT1 = ""
    GLabelJT2 = ""
    
End Sub

Private Sub lstSeq_Click()
    
    If lstSeq.ListIndex = -1 Then Exit Sub
    
    txtDaySeq.Text = lstSeq.List(lstSeq.ListIndex)
    
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

