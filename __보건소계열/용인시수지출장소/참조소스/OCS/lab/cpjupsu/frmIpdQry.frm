VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmIpdQry 
   BackColor       =   &H00C0C0C0&
   Caption         =   "병동환자조회조건"
   ClientHeight    =   3570
   ClientLeft      =   2745
   ClientTop       =   3480
   ClientWidth     =   3570
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
   ScaleHeight     =   3570
   ScaleWidth      =   3570
   Begin Threed.SSCommand cmdWarCclear 
      Height          =   330
      Left            =   3195
      TabIndex        =   7
      Top             =   1215
      Width           =   240
      _Version        =   65536
      _ExtentX        =   423
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "c"
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   735
      Left            =   135
      TabIndex        =   4
      Top             =   90
      Width           =   3300
      _Version        =   65536
      _ExtentX        =   5821
      _ExtentY        =   1296
      _StockProps     =   15
      Caption         =   " 일자조건:From/To="
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
      Begin MSComCtl2.DTPicker dtFrDate 
         Height          =   330
         Left            =   315
         TabIndex        =   5
         Top             =   270
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24576003
         CurrentDate     =   36383
      End
      Begin MSComCtl2.DTPicker dtToDate 
         Height          =   330
         Left            =   1665
         TabIndex        =   6
         Top             =   270
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24576003
         CurrentDate     =   36383
      End
   End
   Begin VB.TextBox txtWardN 
      BackColor       =   &H00C0E0FF&
      Height          =   330
      Left            =   1845
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1215
      Width           =   1320
   End
   Begin VB.TextBox txtWardC 
      BackColor       =   &H00C0E0FF&
      Height          =   330
      Left            =   1845
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   900
      Width           =   465
   End
   Begin VB.ListBox lstWard 
      Height          =   2400
      Left            =   135
      TabIndex        =   0
      Top             =   900
      Width           =   1680
   End
   Begin MSForms.CommandButton cmdExit 
      Height          =   600
      Left            =   1845
      TabIndex        =   8
      Top             =   2700
      Width           =   1545
      Caption         =   "   Exit"
      PicturePosition =   327683
      Size            =   "2725;1058"
      Picture         =   "frmIpdQry.frx":0000
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdQryOk 
      Height          =   600
      Left            =   1845
      TabIndex        =   2
      Top             =   2115
      Width           =   1545
      Caption         =   "조회확인"
      PicturePosition =   327683
      Size            =   "2725;1058"
      Picture         =   "frmIpdQry.frx":08DA
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmIpdQry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
    
End Sub

Private Sub cmdQryOK_Click()
    Dim sFrJeobsuDt         As String
    Dim sToJeobsuDt         As String
    Dim sCompare            As String
    
    sFrJeobsuDt = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToJeobsuDt = Format(dtToDate.Value, "yyyy-MM-dd")
    
    
    GoSub Spread_ssOrder_Clear
    GoSub DataSelect_Main_Process
    
    If frmMain.ssOrder.DataRowCnt = 0 Then
        MsgBox "해당조건에 맞는 환자가 없습니다!..(조건확인)", _
                vbCritical, _
               "Patient Not Found"
    Else
        DoEvents
        Unload Me
    End If
    
    Exit Sub
    
    
DataSelect_Main_Process:
    
    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INX_PATIENT0) */"
    
    strSql = ""
    strSql = strSql & " SELECT a.*, a.RowID OrderRowID,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') Jeobsudt,"
    strSql = strSql & "        TO_CHAR(a.Indate,   'YYYY-MM-DD') Indate,  "
    strSql = strSql & "        TO_CHAR(a.OrderDt,  'YYYY-MM-DD') Orderdt, "
    strSql = strSql & "        TO_CHAR(a.CollDate, 'YYYY-MM-DD') CollDate,"
    strSql = strSql & "        b.Sname, c.Codenm SLname, e.Codenm Samplename, f.Drname"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Order   a, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PATIENT  b, "
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Specode c, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_ROOM     d, "
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Sample  e, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR   f  "
    strSql = strSql & " WHERE  a.JeobsuDt >=  TO_DATE('" & sFrJeobsuDt & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.JeobsuDt <=  TO_DATE('" & sToJeobsuDt & "','yyyy-MM-dd')"
    strSql = strSql & " AND   (a.JeobsuYn  = ' ' Or a.JeobsuYn IS NULL)"
'C      strSql = strSql & " AND    a.SLipno1  <   61"
    strSql = strSql & " AND    a.SLipno1  <   90"
    strSql = strSql & " AND    a.Gbio      = 'I'"
    strSql = strSql & " AND    a.Ptno      = b.Ptno(+)"
    strSql = strSql & " AND    c.Codegu    = '12'"
    strSql = strSql & " AND    TO_NUMBER(c.Codeky)  = a.SLipno1"
    strSql = strSql & " AND    a.GeomchCd  =  e.Code(+)"
    strSql = strSql & " AND    a.Drcode    =  f.Drcode(+)"
    strSql = strSql & " AND    a.RoomCode  =  d.RoomCode"
    
    If Trim(txtWardC.Text) <> "" Then
        strSql = strSql & " AND    d.WardCode  =  '" & Trim(txtWardC.Text) & "'"
    End If
        
    strSql = strSql & " ORDER  BY a.Jeobsudt, a.RoomCode, a.Ptno, a.SLipno1, a.DeptCode"
    
    frmMain.ssOrder.MaxRows = 0
    If False = adoSetOpen(strSql, adoSet) Then Return
    frmMain.ssOrder.MaxRows = adoSet.RecordCount
    
    DoEvents
    Do Until adoSet.EOF
        frmMain.ssOrder.Row = frmMain.ssOrder.DataRowCnt + 1
        frmMain.ssOrder.Col = 2:  frmMain.ssOrder.Text = adoSet.Fields("JeobsuDt").Value & "" & _
                                         adoSet.Fields("Ptno").Value & ""

        frmMain.ssOrder.Col = 2
        If sCompare <> frmMain.ssOrder.Text Then
            frmMain.ssOrder.Col = 4:  frmMain.ssOrder.Text = adoSet.Fields("Jeobsudt").Value & ""
            frmMain.ssOrder.Col = 5:  frmMain.ssOrder.Text = adoSet.Fields("Ptno").Value & ""
            frmMain.ssOrder.Col = 6:  frmMain.ssOrder.Text = adoSet.Fields("Sname").Value & ""
            frmMain.ssOrder.Col = 7:  frmMain.ssOrder.Text = adoSet.Fields("Sex").Value & ""
            frmMain.ssOrder.Col = 8:  frmMain.ssOrder.Text = adoSet.Fields("AgeYY").Value & ""
            frmMain.ssOrder.Col = 9:  frmMain.ssOrder.Text = adoSet.Fields("AgeMM").Value & ""
        Else
            frmMain.ssOrder.Col = 1:   frmMain.ssOrder.CellType = CellTypeStaticText
            frmMain.ssOrder.BackColor = RGB(254, 255, 240)
        End If
        
        frmMain.ssOrder.Col = 3:   frmMain.ssOrder.Text = adoSet.Fields("OrderRowID").Value & ""
        frmMain.ssOrder.Col = 10:  frmMain.ssOrder.Text = adoSet.Fields("SLipno1").Value & ""
        frmMain.ssOrder.Col = 11:  frmMain.ssOrder.Text = adoSet.Fields("SLname").Value & ""
                
        frmMain.ssOrder.Col = 12: frmMain.ssOrder.Text = adoSet.Fields("Itemcd").Value & ""
        
        If IsRoutineCode(adoSet.Fields("ItemCd").Value & "") Then
            frmMain.ssOrder.Col = 13: frmMain.ssOrder.Text = Get_RoutineName(adoSet.Fields("ItemCD").Value & "")
        Else
            frmMain.ssOrder.Col = 13: frmMain.ssOrder.Text = Get_ItemName(adoSet.Fields("ItemCD").Value & "")
        End If
                
        
        
        frmMain.ssOrder.Col = 14:  frmMain.ssOrder.Text = Format(adoSet.Fields("JeobsuT1").Value, "00") & ":" & _
                                         Format(adoSet.Fields("JeobsuT2").Value, "00")
        
        frmMain.ssOrder.Col = 15: frmMain.ssOrder.Text = adoSet.Fields("Indate").Value & ""
        frmMain.ssOrder.Col = 16: frmMain.ssOrder.Text = adoSet.Fields("RoomCode").Value & ""
        frmMain.ssOrder.Col = 17: frmMain.ssOrder.Text = adoSet.Fields("DeptCode").Value & ""
        frmMain.ssOrder.Col = 18: frmMain.ssOrder.Text = adoSet.Fields("Gbio").Value & ""
        frmMain.ssOrder.Col = 19: frmMain.ssOrder.Text = adoSet.Fields("Bi").Value & ""
        frmMain.ssOrder.Col = 20: frmMain.ssOrder.Text = adoSet.Fields("GbER").Value & ""
        frmMain.ssOrder.Col = 21: frmMain.ssOrder.Value = True
        
        frmMain.ssOrder.Col = 22: frmMain.ssOrder.Text = adoSet.Fields("GeomchCD").Value & ""
        frmMain.ssOrder.Col = 23: frmMain.ssOrder.Text = adoSet.Fields("Samplename").Value & ""
        
        frmMain.ssOrder.Col = 24: frmMain.ssOrder.Text = adoSet.Fields("GeomsaGu").Value & ""
        frmMain.ssOrder.Col = 25: frmMain.ssOrder.Text = adoSet.Fields("OrderDt").Value & ""
        frmMain.ssOrder.Col = 26: frmMain.ssOrder.Text = adoSet.Fields("OrderNo").Value & ""
        frmMain.ssOrder.Col = 27: frmMain.ssOrder.Text = adoSet.Fields("OrderCD").Value & ""
        frmMain.ssOrder.Col = 28: frmMain.ssOrder.Text = adoSet.Fields("Quantity").Value & ""
        frmMain.ssOrder.Col = 29: frmMain.ssOrder.Text = adoSet.Fields("CmDoctor").Value & ""
        frmMain.ssOrder.Col = 30: frmMain.ssOrder.Text = adoSet.Fields("DrCode").Value & ""
        frmMain.ssOrder.Col = 31: frmMain.ssOrder.Text = adoSet.Fields("Drname").Value & ""
        frmMain.ssOrder.Col = 32: frmMain.ssOrder.Text = adoSet.Fields("JeobsuYn").Value & ""
        frmMain.ssOrder.Col = 33: frmMain.ssOrder.Text = adoSet.Fields("Gbinfo").Value & ""
        
        
        frmMain.ssOrder.Col = 34: frmMain.ssOrder.Text = adoSet.Fields("CollDate").Value & ""
        frmMain.ssOrder.Col = 35: frmMain.ssOrder.Text = adoSet.Fields("CollHH").Value & ""
        frmMain.ssOrder.Col = 36: frmMain.ssOrder.Text = adoSet.Fields("CollMM").Value & ""
        frmMain.ssOrder.Col = 37: frmMain.ssOrder.Text = adoSet.Fields("Jeobsu_Lab").Value & ""
        
        sCompare = adoSet.Fields("JeobsuDt").Value & "" & _
                   adoSet.Fields("Ptno").Value & ""
        
        adoSet.MoveNext
    Loop
    
    
    Call adoSetClose(adoSet)
    Return
    
Spread_ssOrder_Clear:
    frmMain.ssOrder.ReDraw = False
    frmMain.ssOrder.MaxRows = 0
    frmMain.ssOrder.MaxRows = 20
    frmMain.ssOrder.RowHeight(-1) = 11.5
    frmMain.ssOrder.ReDraw = True
    
    frmMain.ssEnrol.ReDraw = False
    frmMain.ssEnrol.MaxRows = 0
    frmMain.ssEnrol.MaxRows = 500
    frmMain.ssEnrol.RowHeight(-1) = 11
    frmMain.ssEnrol.ReDraw = True
    
    frmMain.sprLabno.ReDraw = False
    frmMain.sprLabno.MaxRows = 0
    frmMain.sprLabno.MaxRows = 20
    frmMain.sprLabno.RowHeight(-1) = 11
    frmMain.sprLabno.ReDraw = True
    
    
    For i = 0 To frmMain.Count - 1
        If TypeOf frmMain.Controls(i) Is VB.TextBox Then frmMain.Controls(i).Text = ""
    Next

    Return

End Sub

Private Sub cmdWarCclear_Click()
    
    txtWardC.Text = ""
    txtWardN.Text = ""
    
End Sub

Private Sub Form_Load()

    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    GoSub Get_Date_Setting
    GoSub Get_Ward_Data
    Exit Sub
    
    
    
    
Get_Date_Setting:
    dtFrDate.Value = Dual_Date_Cal_Get("yyyy-MM-dd", -1)
    dtToDate.Value = Dual_Date_Get("yyyy-MM-dd")
    Return
    
Get_Ward_Data:
    Dim sWardC  As String * 4
    
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TW_MIS_PMPA.TWBAS_WARD"
    strSql = strSql & " ORDER  BY WardCode"
    If False = adoSetOpen(strSql, adoSet) Then Return
    Do Until adoSet.EOF
        sWardC = adoSet.Fields("WardCode").Value & ""
        lstWard.AddItem sWardC & Trim(adoSet.Fields("WardName").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return

End Sub

Private Sub lstWard_Click()
    If lstWard.ListIndex = -1 Then Exit Sub
    
    txtWardC.Text = Left(lstWard.Text, 4)
    txtWardN.Text = Mid(lstWard.Text, 5, Len(lstWard.Text) - 4)
    
    
End Sub

