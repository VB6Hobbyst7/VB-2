VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmResult 
   Caption         =   "결과조회"
   ClientHeight    =   8085
   ClientLeft      =   2340
   ClientTop       =   1125
   ClientWidth     =   11580
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
   ScaleHeight     =   8085
   ScaleWidth      =   11580
   WindowState     =   2  '최대화
   Begin FPSpreadADO.fpSpread ssDtList 
      Height          =   6180
      Left            =   4080
      TabIndex        =   3
      Top             =   765
      Width           =   7710
      _Version        =   196608
      _ExtentX        =   13600
      _ExtentY        =   10901
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
      MaxCols         =   10
      ScrollBars      =   2
      SpreadDesigner  =   "frmResult.frx":0000
      Appearance      =   1
   End
   Begin VB.TextBox txtGeomsaJa 
      BackColor       =   &H00C00000&
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   10350
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6975
      Width           =   1410
   End
   Begin VB.TextBox txtGeomsaCm 
      BackColor       =   &H00FFC0C0&
      Height          =   645
      Left            =   4095
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6975
      Width           =   6225
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   330
      Index           =   0
      Left            =   135
      TabIndex        =   9
      Top             =   765
      Width           =   2085
      _Version        =   65536
      _ExtentX        =   3678
      _ExtentY        =   582
      _StockProps     =   15
      Caption         =   "검사종목"
      ForeColor       =   0
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
      BorderWidth     =   1
      BevelInner      =   2
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   555
      Left            =   180
      TabIndex        =   4
      Top             =   90
      Width           =   11400
      _Version        =   65536
      _ExtentX        =   20108
      _ExtentY        =   979
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
      Begin VB.TextBox txtRoom 
         Enabled         =   0   'False
         Height          =   285
         Left            =   10530
         TabIndex        =   20
         Top             =   135
         Width           =   600
      End
      Begin VB.TextBox txtDrname 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   8370
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   135
         Width           =   1140
      End
      Begin VB.TextBox txtDeptNameK 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   6480
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   135
         Width           =   1140
      End
      Begin Threed.SSCommand cmdQrySname 
         Height          =   285
         Left            =   3780
         TabIndex        =   11
         Top             =   135
         Width           =   240
         _Version        =   65536
         _ExtentX        =   423
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "&S"
      End
      Begin VB.TextBox txtAge 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   5265
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "999"
         Top             =   135
         Width           =   420
      End
      Begin VB.TextBox txtSex 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   5040
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "M"
         Top             =   135
         Width           =   240
      End
      Begin VB.TextBox txtSname 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   2700
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "홍길동아가"
         Top             =   135
         Width           =   1095
      End
      Begin VB.TextBox txtPtno 
         Height          =   285
         Left            =   990
         TabIndex        =   0
         Text            =   "12345678"
         Top             =   135
         Width           =   960
      End
      Begin VB.Label Label6 
         Caption         =   "병실?"
         Height          =   195
         Left            =   9990
         TabIndex        =   21
         Top             =   180
         Width           =   510
      End
      Begin VB.Label Label5 
         Caption         =   "담당의"
         Height          =   195
         Left            =   7740
         TabIndex        =   16
         Top             =   180
         Width           =   600
      End
      Begin VB.Label Label4 
         Caption         =   "진료과"
         Height          =   195
         Left            =   5850
         TabIndex        =   14
         Top             =   180
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "성별/나이"
         Height          =   240
         Left            =   4185
         TabIndex        =   13
         Top             =   180
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "환자명"
         Height          =   195
         Left            =   2115
         TabIndex        =   12
         Top             =   180
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "병록번호"
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   180
         Width           =   825
      End
   End
   Begin FPSpreadADO.fpSpread ssDayList 
      Height          =   5145
      Left            =   2295
      TabIndex        =   2
      Top             =   1080
      Width           =   1770
      _Version        =   196608
      _ExtentX        =   3122
      _ExtentY        =   9075
      _StockProps     =   64
      DisplayColHeaders=   0   'False
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
      MaxCols         =   10
      ScrollBars      =   0
      SpreadDesigner  =   "frmResult.frx":19C0
      Appearance      =   1
   End
   Begin FPSpreadADO.fpSpread ssList 
      Height          =   5145
      Left            =   135
      TabIndex        =   1
      Top             =   1080
      Width           =   2085
      _Version        =   196608
      _ExtentX        =   3678
      _ExtentY        =   9075
      _StockProps     =   64
      DisplayColHeaders=   0   'False
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
      MaxRows         =   50
      ScrollBars      =   0
      SpreadDesigner  =   "frmResult.frx":561F
      UserResize      =   1
      Appearance      =   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   330
      Index           =   1
      Left            =   2295
      TabIndex        =   10
      Top             =   765
      Width           =   1770
      _Version        =   65536
      _ExtentX        =   3122
      _ExtentY        =   582
      _StockProps     =   15
      Caption         =   "일자별List"
      ForeColor       =   0
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
      BorderWidth     =   1
      BevelInner      =   2
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   1800
      Picture         =   "frmResult.frx":73F8
      Stretch         =   -1  'True
      Top             =   6480
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQrySname_Click()
    
    hWndReturn = txtPtno.hwnd
    frmQrySname.Show vbModal
    
    If Trim(txtPtno.Text) <> "" Then
        Call txtPtno_KeyPress(13)
    End If
    
End Sub

Private Sub Form_Activate()

    Me.WindowState = vbMaximized
    
End Sub

Private Sub Form_Load()
    
    For i = 0 To Me.Count - 1
        If TypeOf Me.Controls(i) Is VB.TextBox Then
            Me.Controls(i).Text = ""
        End If
    Next
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

Private Sub ssDayList_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim sSLipno1    As String
    Dim sJeobsuDt   As String
    Dim sHiCham     As String
    Dim sLoCham     As String
    Dim sResult(1 To 5) As String
    Dim sRcode(1 To 5) As String
    
    
    Dim sSensSL1    As String
    Dim sSensSL2    As String
    Dim sSensItemCd As String
    
    
    ssDayList.Row = Row
    ssDayList.Col = 2
    If (ssDayList.Text) = "" Then Exit Sub
    
    If Col = 1 Then
        GoSub Marking_Set
        ssDayList.Row = Row
        ssDayList.Col = 2: sJeobsuDt = ssDayList.Text
        ssDayList.Col = 3: sSLipno1 = ssDayList.Text
        If Trim(sSLipno1) = "15" Then        '골수검사일경우
            ssDtList.ColWidth(2) = 21
            ssDtList.ColWidth(3) = 41
            GoSub Get_General_Sub_BMData
        Else
            ssDtList.ColWidth(2) = 27.88
            ssDtList.ColWidth(3) = 17.5
            GoSub Get_General_Sub_Data
        End If
        
    End If
    Exit Sub
    
    
    
Marking_Set:
    ssDayList.Col = 1: ssDayList.Col2 = 1
    ssDayList.Row = 1: ssDayList.Row2 = ssDayList.DataRowCnt
    ssDayList.BlockMode = True
    'ssDayList.TypeButtonColor = RGB(192, 192, 192)
    ssDayList.TypeButtonText = ""
    ssDayList.BlockMode = False


    ssDayList.Col = 1
    ssDayList.Row = Row
    'ssDayList.TypeButtonColor = RGB(0, 255, 0)
    ssDayList.TypeButtonText = "▶"
    
    Return
    
Get_General_Sub_BMData:
    Dim sDispResult1        As String * 5
    
    strSql = ""
    strSql = strSql & " SELECT a.*,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDt,'YYYY-MM-DD') JeobsuDt, "
    strSql = strSql & "        b.Itemnm, b.ResultW, c.GeomsaCm, c.GeomsaJa"
    strSql = strSql & " FROM   TWEXAM_General_Sub a,"
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_itemML      b,"
    strSql = strSql & "        TWEXAM_General     c "
    strSql = strSql & " WHERE  a.Ptno     =  '" & txtPtno.Text & "'"
    strSql = strSql & " AND    a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.SLipno1  = " & Val(sSLipno1)
    strSql = strSql & " AND    a.ItemCd   = b.Codeky(+)"
    strSql = strSql & " AND    a.JeobsuDt = c.JeobsuDt(+)"
    strSql = strSql & " AND    a.SLipno1  = c.SLipno1(+)"
    strSql = strSql & " AND    a.SLipno2  = c.SLipno2(+)"
    strSql = strSql & " ORDER  BY a.ItemCd"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    txtGeomsaCm.Text = ""
    txtGeomsaJa.Text = ""
    Call Spread_Set_Clear(ssDtList)
    ssDtList.RowHeight(-1) = 9.55
    
    Do Until adoSet.EOF
    
        ssDtList.Row = ssDtList.DataRowCnt + 1
        ssDtList.Col = 1: ssDtList.Text = adoSet.Fields("ItemCd").Value & ""
        ssDtList.Col = 2: ssDtList.Text = adoSet.Fields("ItemNm").Value & ""
        
        ssDtList.Col = 3
        Select Case Trim(adoSet.Fields("ItemCD").Value & "")
            Case "1505011": ssDtList.Text = Trim(adoSet.Fields("Chamgo").Value & "")  '골수검사 PBSmear
            Case "1505012": ssDtList.Text = Trim(adoSet.Fields("Chamgo").Value & "")  '골수검사 Aspiration
            Case "15050201" To "15050218"                                             '골수검사 Diffcount
                sDispResult1 = Trim(adoSet.Fields("Result1").Value & "")
                If Trim(sDispResult1) <> "" Then
                    ssDtList.Text = sDispResult1 & "%"
                End If
        End Select
        ssDtList.RowHeight(ssDtList.Row) = ssDtList.MaxTextRowHeight(ssDtList.Row)
        If Trim(txtGeomsaCm.Text) = "" Then
            txtGeomsaCm.Text = Trim(adoSet.Fields("GeomsaCm").Value & "")
        End If
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    If Trim(txtGeomsaJa.Text) <> "" Then
        GoSub Get_PassName_Data
    End If
    
    Return
    
    
    
    
    
Get_General_Sub_Data:
    strSql = ""
    strSql = strSql & " SELECT a.*,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDt,'YYYY-MM-DD') JeobsuDt, "
    strSql = strSql & "        b.Itemnm, b.ResultW, c.GeomsaCm, c.GeomsaJa"
    strSql = strSql & " FROM   TWEXAM_General_Sub a,"
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_itemML      b,"
    strSql = strSql & "        TWEXAM_General     c "
    strSql = strSql & " WHERE  a.Ptno     =  '" & txtPtno.Text & "'"
    strSql = strSql & " AND    a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.SLipno1  = " & Val(sSLipno1)
    strSql = strSql & " AND    a.ItemCd   = b.Codeky(+)"
    strSql = strSql & " AND    a.JeobsuDt = c.JeobsuDt(+)"
    strSql = strSql & " AND    a.SLipno1  = c.SLipno1(+)"
    strSql = strSql & " AND    a.SLipno2  = c.SLipno2(+)"
    strSql = strSql & " ORDER  BY a.ItemCd"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    txtGeomsaCm.Text = ""
    txtGeomsaJa.Text = ""
    Call Spread_Set_Clear(ssDtList)
    ssDtList.RowHeight(-1) = 9.55
    
    
    Do Until adoSet.EOF
        ssDtList.Row = ssDtList.DataRowCnt + 1
        ssDtList.Col = 1: ssDtList.Text = adoSet.Fields("ItemCd").Value & ""
                          sSensItemCd = adoSet.Fields("ItemCd").Value & ""
        ssDtList.Col = 2: ssDtList.Text = adoSet.Fields("ItemNm").Value & ""
        
        For i = 1 To 5
            sResult(i) = Trim(adoSet.Fields("Result" & i).Value & "")
            sRcode(i) = Trim(adoSet.Fields("Rcode" & i).Value & "")
        Next
        ssDtList.Col = 6: ssDtList.Text = adoSet.Fields("ResultW").Value & ""
        
        If Trim(ssDtList.Text) = "S" Then 'Senstivity
            For i = 1 To 5
                If Trim(sRcode(i)) <> "" Or Trim(sResult(i)) <> "" Then
                    ssDtList.Row = ssDtList.DataRowCnt + 1
                    ssDtList.Col = 2: ssDtList.Text = "    (" & i & ")." & sRcode(i) & vbCrLf & "        " & sResult(i)
                    GoSub Senstivity_Result_Get
                End If
            Next
        Else
            GoSub Get_RefData
            ssDtList.Col = 3: ssDtList.Text = convResultFormat(sResult(1))
            ssDtList.Col = 4: ssDtList.Text = convResultFormat(sLoCham)
            ssDtList.Col = 5: ssDtList.Text = convResultFormat(sHiCham)
        End If
        ssDtList.Col = 7: ssDtList.Text = adoSet.Fields("JeobsuDt").Value & ""
        ssDtList.Col = 8: ssDtList.Text = adoSet.Fields("SLipno1").Value & ""
                          sSensSL1 = adoSet.Fields("SLipno1").Value & ""
        ssDtList.Col = 9: ssDtList.Text = adoSet.Fields("SLipno2").Value & ""
                          sSensSL2 = adoSet.Fields("SLipno2").Value & ""
        txtGeomsaJa.Text = Trim(adoSet.Fields("GeomsaJa").Value & "")
        If Trim(txtGeomsaCm.Text) = "" Then
            txtGeomsaCm.Text = Trim(adoSet.Fields("GeomsaCm").Value & "")
        End If
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    If Trim(txtGeomsaJa.Text) <> "" Then
        GoSub Get_PassName_Data: End If
    
    Return
    
Get_PassName_Data:
    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TWBAS_PASS INDEX_PASS0) */"
    
    strSql = ""
    strSql = strSql & " SELECT Name"
    strSql = strSql & " FROM   TW_MIS_PMPA.TWBAS_Pass"
    strSql = strSql & " WHERE  idNumber  =  '" & Trim(txtGeomsaJa.Text) & "'"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    txtGeomsaJa.Text = adoSet.Fields("Name").Value & ""
    Call adoSetClose(adoSet)
    
    Return
    
Senstivity_Result_Get:
    Dim adoSens     As ADODB.Recordset
    Dim iLastData   As Integer
    
    strSql = ""
    strSql = strSql & " SELECT a.*, b.Codenm"
    strSql = strSql & " FROM   TWEXAM_Sens     a,"
    strSql = strSql & "        TWEXAM_AntiList b "
    strSql = strSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.SLipno1  =  " & Val(sSensSL1)
    strSql = strSql & " AND    a.SLipno2  =  " & Val(sSensSL2)
    strSql = strSql & " AND    a.OraCod   = '" & Trim(sRcode(i)) & "'"
    strSql = strSql & " AND    a.ItemCD   = '" & Trim(sSensItemCd) & "'"
    strSql = strSql & " AND    a.YakCod   = b.Codeky(+)"
    If False = adoSetOpen(strSql, adoSens) Then Return
    
    ssDtList.ReDraw = False
    iLastData = 1
    Do Until adoSens.EOF
        If iLastData = adoSens.RecordCount Then
            ssDtList.Col = 3: ssDtList.Text = ssDtList.Text & _
                                              Trim(adoSens.Fields("Codenm").Value & "") & "(" & _
                                              Trim(adoSens.Fields("Sens").Value) & ")"
        Else
            ssDtList.Col = 3: ssDtList.Text = ssDtList.Text & _
                                              Trim(adoSens.Fields("Codenm").Value & "") & "(" & _
                                              Trim(adoSens.Fields("Sens").Value) & ")" & vbCrLf
        End If
        ssDtList.RowHeight(ssDtList.Row) = ssDtList.MaxTextRowHeight(ssDtList.Row)
        adoSens.MoveNext: iLastData = iLastData + 1
    Loop
    Call adoSetClose(adoSens)
    ssDtList.ReDraw = True
    Return
    

Get_RefData:
    Dim adoRef      As ADODB.Recordset
    
    sHiCham = ""
    sLoCham = ""
    
    strSql = ""
    strSql = strSql & " SELECT * "
    strSql = strSql & " FROM   TWEXAM_REFDATA"
    strSql = strSql & " WHERE  ITEMCODE  = '" & Trim(adoSet.Fields("ItemCd").Value & "") & "'"
    strSql = strSql & " AND    APPDATE   =     (SELECT MAX(APPDATE)"
    strSql = strSql & "                         FROM   TWEXAM_REFDATA"
    strSql = strSql & "                         WHERE  ITEMCODE = '" & Trim(adoSet.Fields("ItemCd").Value & "") & "'"
    strSql = strSql & "                         AND    AGEMIN  <=  " & Val(txtAge.Text)
    strSql = strSql & "                         AND    AGEMAX  >=  " & Val(txtAge.Text) & ")"
    strSql = strSql & " AND    AGEMIN   <=  " & Val(txtAge.Text)
    strSql = strSql & " AND    AGEMAX   >=  " & Val(txtAge.Text)
    
    If False = adoSetOpen(strSql, adoRef) Then Return
    
    If Trim(txtSex.Text) = "M" Then
        sLoCham = Trim(adoRef.Fields("M_MIN").Value & "")
        sHiCham = Trim(adoRef.Fields("M_MAX").Value & ""): End If
    If Trim(txtSex.Text) = "F" Then
        sLoCham = Trim(adoRef.Fields("F_MIN").Value & "")
        sHiCham = Trim(adoRef.Fields("F_MAX").Value & ""): End If
    Call adoSetClose(adoRef)

    
    Return

End Sub

Private Sub ssList_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim sPtno       As String
    Dim sSLipno1    As String
    
    ssList.Row = Row
    ssList.Col = 2
    If (ssList.Text) = "" Then Exit Sub

    If Col = 1 Then
        GoSub SPREAD_CLEAR_SUB
        GoSub Hand_Flag_Set
        ssList.Row = Row
        ssList.Col = 3: sSLipno1 = ssList.Text
        GoSub GET_Date_SLip
    End If
    Exit Sub
    
    

Hand_Flag_Set:
    'Col 1 의 Hand Picture Reset
    ssList.Row = 1:  ssList.Row2 = ssList.DataRowCnt
    ssList.Col = 1:  ssList.Col2 = 1
    ssList.BlockMode = True
    ssList.TypeButtonPicture = LoadPicture("")
    ssList.BlockMode = False
    
    'Text Color Reset
    ssList.Row = 1
    ssList.Row2 = ssList.MaxRows
    ssList.Col = 2
    ssList.Col2 = ssList.MaxCols
    ssList.BlockMode = True
    ssList.ForeColor = RGB(0, 0, 0)
    ssList.BlockMode = False
    
    'Select Row Hand-Picture Set
    ssList.Row = Row
    ssList.Col = 1
    If ssList.CellType = CellTypeButton Then
        ssList.TypeButtonPicture = Image1.Picture
        'ssList.TypeButtonPicture = LoadPicture("c:\twhis\src60\ocs\lab\data\fingerr.bmp")
        ssList.Row = Row
        ssList.Row2 = Row
        ssList.Col = 2
        ssList.Col2 = ssList.MaxCols
        ssList.BlockMode = True
        ssList.ForeColor = RGB(192, 0, 220)
        ssList.BlockMode = False
    End If
    Return
    
    

SPREAD_CLEAR_SUB:
    Call Spread_Set_Clear(ssDayList)
    Call Spread_Set_Clear(ssDtList)
    txtGeomsaJa.Text = ""
    txtGeomsaCm.Text = ""
    Return
    
    
GET_Date_SLip:
    strSql = ""
    strSql = strSql & " SELECT DISTINCT TO_CHAR(JeobsuDt, 'YYYY-MM-DD') JeobsuDt"
    strSql = strSql & " FROM   TWEXAM_General"
    strSql = strSql & " WHERE  Ptno    =  '" & txtPtno.Text & "'"
    strSql = strSql & " AND    SLipno1 = " & Val(sSLipno1)
    strSql = strSql & " ORDER  BY JeobsuDt DESC"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Call Spread_Set_Clear(ssDayList)
    
    Do Until adoSet.EOF
        ssDayList.Row = ssDayList.DataRowCnt + 1
        ssDayList.Col = 2: ssDayList.Text = adoSet.Fields("JeobsuDt").Value & ""
        ssDayList.Col = 3: ssDayList.Text = sSLipno1
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return
    
End Sub

Private Sub txtPtno_GotFocus()
    
    txtPtno.SelStart = 0
    txtPtno.SelLength = Len(txtPtno.Text)
    
End Sub

Private Sub txtPtno_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txtPtno.Text = UCase(txtPtno.Text)
        txtPtno.Text = Format(txtPtno.Text, "00000000")
        If Trim(txtPtno.Text) = "" Then Exit Sub
        GoSub TEXTBox_Clear
        GoSub SPREAD_ALL_CLEAR
        GoSub Get_HJ_Information
        GoSub Get_First_Process
    End If
    Exit Sub
    
'/------------------------------------------------------------
TEXTBox_Clear:
    txtPtno.Tag = txtPtno.Text
    
    For i = 0 To Me.Count - 1
        If TypeOf Me.Controls(i) Is VB.TextBox Then
            Me.Controls(i).Text = ""
        End If
    Next
    
    txtPtno.Text = txtPtno.Tag
    
    Return
    
    
SPREAD_ALL_CLEAR:
    Call Spread_Set_Clear(ssList)
    Call Spread_Set_Clear(ssDayList)
    Call Spread_Set_Clear(ssDtList)
    Return
    


Get_HJ_Information:
    'TWEXAM_IDNOMST
    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_DEPT INDEX_DEPT0) */"
    
    strSql = ""
    strSql = strSql & " SELECT a.Sname, a.Sex, a.AgeYY,"
    strSql = strSql & "        TO_CHAR(a.BirthDay, 'YYYY-MM-DD') BirthDay,"
    strSql = strSql & "        b.DeptnameK, c.Drname"
    strSql = strSql & " FROM   TWEXAM_IDNOMST a,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT     b,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR   c "
    strSql = strSql & " WHERE  a.Ptno     = '" & txtPtno.Text & "'"
    strSql = strSql & " AND    a.DeptCode = b.DeptCode(+)"
    strSql = strSql & " AND    a.Drcode   = c.Drcode(+)"
    If adoSetOpen(strSql, adoSet) Then
        txtSname.Text = Trim(adoSet.Fields("Sname").Value & "")
        txtSex.Text = Trim(adoSet.Fields("Sex").Value & "")
        txtAge.Text = Trim(adoSet.Fields("Ageyy").Value & "")
        txtDeptNameK.Text = Trim(adoSet.Fields("DeptNamek").Value & "")
        txtDrname.Text = Trim(adoSet.Fields("Drname").Value & "")
        Call adoSetClose(adoSet)
    Else
        'HJ_MASTER
        'strSql = ""
        'strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INDEX_PATIENT0) */"
        
        strSql = ""
        strSql = strSql & " SELECT a.Sname, a.Sex, a.Jumin1, a.Jumin2,"
        strSql = strSql & "        b.DeptNameK, c.Drname"
        strSql = strSql & " FROM   TW_MIS_PMPA.TWBAS_PATIENT a,"
        strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT    b,"
        strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR  c "
        strSql = strSql & " WHERE  a.Ptno     =  '" & txtPtno.Text & "'"
        strSql = strSql & " AND    a.DeptCode = b.DeptCode(+)"
        strSql = strSql & " AND    a.Drcode   = c.Drcode(+)"
        
        If False = adoSetOpen(strSql, adoSet) Then Return
        
        txtSname.Text = Trim(adoSet.Fields("Sname").Value & "")
        txtSex.Text = Trim(adoSet.Fields("Sex").Value & "")
        txtAge.Text = SetAge_Check(adoSet.Fields("Jumin1").Value & "", adoSet.Fields("Jumin2").Value & "")
        txtDeptNameK.Text = Trim(adoSet.Fields("DeptNamek").Value & "")
        txtDrname.Text = Trim(adoSet.Fields("Drname").Value & "")
        Call adoSetClose(adoSet)
    End If
    
   '재원환자일 경우 병실코드를 Display 시킨다.
    If IsAdmission(txtPtno.Text) Then
        'strSql = ""
        'strSql = strSql & " SELECT /*+ INDEX (TWIpd_Master INDEX_IPDMST2) */"
        
        strSql = ""
        strSql = strSql & " SELECT RoomCode "
        strSql = strSql & " FROM   TW_MIS_PMPA.TWIPD_Master "
        strSql = strSql & " WHERE  Ptno = '" & txtPtno.Text & "'"
        If adoSetOpen(strSql, adoSet) Then
            txtRoom.Text = Trim(adoSet.Fields("RoomCode").Value & "")
            Call adoSetClose(adoSet)
        End If
    End If
    
    Return



Get_First_Process:
    strSql = ""
    strSql = strSql & " SELECT a.SLipno1, b.Codenm"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_GENERAL a,"
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Specode b "
    strSql = strSql & " WHERE  a.Ptno    =  '" & txtPtno.Text & "'"
    strSql = strSql & " AND    b.Codeky  = a.SLipno1"                'Item Code
    strSql = strSql & " AND    b.Codegu  = '12'"                     'Special Code
    strSql = strSql & " GROUP  BY a.SLipno1, b.Codenm"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Call Spread_Set_Clear(ssList)
    Do Until adoSet.EOF
        ssList.Row = ssList.DataRowCnt + 1
        ssList.Col = 2: ssList.Text = adoSet.Fields("Codenm").Value & ""
        ssList.Col = 3: ssList.Text = adoSet.Fields("SLipno1").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return
    
End Sub
