VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmRpt 
   Caption         =   "Report Ãâ·Â"
   ClientHeight    =   7785
   ClientLeft      =   1800
   ClientTop       =   4965
   ClientWidth     =   11535
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "±¼¸²Ã¼"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7785
   ScaleWidth      =   11535
   WindowState     =   2  'ÃÖ´ëÈ­
   Begin Threed.SSPanel panelChemi 
      Height          =   5865
      Left            =   135
      TabIndex        =   1
      Top             =   1710
      Width           =   11310
      _Version        =   65536
      _ExtentX        =   19950
      _ExtentY        =   10345
      _StockProps     =   15
      Caption         =   "Chemistry"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      RoundedCorners  =   0   'False
      Alignment       =   0
      Begin FPSpreadADO.fpSpread sprChemi 
         Height          =   5145
         Left            =   225
         TabIndex        =   2
         Top             =   405
         Width           =   10905
         _Version        =   196608
         _ExtentX        =   19235
         _ExtentY        =   9075
         _StockProps     =   64
         BackColorStyle  =   1
         ColsFrozen      =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   53
         SpreadDesigner  =   "frmRpt.frx":0000
         UserResize      =   1
         Appearance      =   2
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5175
      Top             =   990
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
            Picture         =   "frmRpt.frx":2227
            Key             =   "Exit"
            Object.Tag             =   "Exit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'À§ ¸ÂÃã
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   635
      ButtonWidth     =   1270
      ButtonHeight    =   582
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
            Description     =   "Exit"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   465
      Left            =   90
      TabIndex        =   3
      Top             =   450
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   820
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      RoundedCorners  =   0   'False
      MouseIcon       =   "frmRpt.frx":2543
      Begin VB.ComboBox cmbSLip 
         Height          =   300
         Left            =   1575
         Style           =   2  'µå·Ó´Ù¿î ¸ñ·Ï
         TabIndex        =   4
         Top             =   90
         Width           =   2670
      End
      Begin VB.Label Label9 
         Caption         =   "°Ë»çÁ¾¸ñ:"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   630
         TabIndex        =   5
         Top             =   135
         Width           =   780
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   90
         Picture         =   "frmRpt.frx":37E5
         Stretch         =   -1  'True
         Top             =   0
         Width           =   480
      End
   End
   Begin MSComCtl2.DTPicker dtTdate 
      Height          =   315
      Left            =   7785
      TabIndex        =   6
      Top             =   630
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   24576003
      CurrentDate     =   36306
   End
   Begin MSComCtl2.DTPicker dtFdate 
      Height          =   315
      Left            =   6390
      TabIndex        =   7
      Top             =   630
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   24576003
      CurrentDate     =   36306
   End
   Begin MSForms.CommandButton cmdChemi 
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   945
      Width           =   1725
      Caption         =   "Á¶È¸È®ÀÎ"
      Size            =   "3043;661"
      FontName        =   "±¼¸²Ã¼"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub cmdChemi_Click()
    Dim sFrDate             As String
    Dim sToDate             As String
    
    
    sFrDate = Format(Me.dtFdate.Value, "yyyy-MM-dd")
    sToDate = Format(Me.dtTdate.Value, "yyyy-MM-dd")
    
    
    
    strSql = ""
    strSql = strSql & "  SELECT  JeobsuDT, Ptno, Sname, Sex, ageyy, DeptCode, RoomCode,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210101', RET, '')) Ca, "
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210102', RET, '')) P, "
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210103', RET, '')) Glu, "
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210104', RET, '')) BUN, "
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210105', RET, '')) Crea, "
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210106', RET, '')) UA, "
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210107', RET, '')) CholT, "
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210108', RET, '')) ProT, "
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210109', RET, '')) Alb,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210110', RET, '')) AkP,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210111', RET, '')) GOT,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210112', RET, '')) GPT,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210113', RET, '')) BiLT,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210114', RET, '')) BiLD,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210115', RET, '')) rGT,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210116', RET, '')) TTT,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210117', RET, '')) CPK,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210118', RET, '')) LDH,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210119', RET, '')) NH3,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210201', RET, '')) Na,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210202', RET, '')) K,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210203', RET, '')) Cl,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210204', RET, '')) Co2,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210301', RET, '')) Amy,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210302', RET, '')) Lip,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210303', RET, '')) AcidPho,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210304', RET, '')) Cholin,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210305', RET, '')) IronS,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210306', RET, '')) UIBC,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210307', RET, '')) OsmoS,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210308', RET, '')) OsmoU,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210310', RET, '')) Mg,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210311', RET, '')) Tg,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210312', RET, '')) HDL,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210313', RET, '')) LDL,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210314', RET, '')) bLipo,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210315', RET, '')) PLipid,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210316', RET, '')) TLipid,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210317', RET, '')) ADA,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210318', RET, '')) Gluef,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210319', RET, '')) Gluef2h,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210401', RET, '')) OGttf,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210402', RET, '')) OGtt3,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210403', RET, '')) OGtt6,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210404', RET, '')) OGtt9,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210405', RET, '')) OGtt12,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210406', RET, '')) OGtt18,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210501', RET, '')) HbA1c,"
    strSql = strSql & "          MAX(DECODE(RTRIM(ITEMCD), '210502', RET, '')) Fruc "
    
    strSql = strSql & "  FROM  ( Select TO_CHAR(a.JeobsuDt,'YYYY-MM-DD') Jeobsudt, a.Ptno, b.Sname, b.Sex, b.ageYY,"
    strSql = strSql & "                 c.DeptCode, c.RoomCode,"
    strSql = strSql & "                 a.ItemCd, RTRIM(SUBSTR(LTRIM(a.RESULT1),1,15)) RET"
    strSql = strSql & "          From   twexam_general_sub  a,"
    strSql = strSql & "                 twexam_idnomst      b,"
    strSql = strSql & "                 twexam_general      c "
    strSql = strSql & "          WHERE  a.SLIPNO1   = 21"
    strSql = strSql & "          And    a.Verify    = 'Y'"
    strSql = strSql & "          And    a.JeobsuDt >= TO_DATE('" & sFrDate & "','YYYY-MM-DD')"
    strSql = strSql & "          And    a.JeobsuDt <= TO_DATE('" & sToDate & "','YYYY-MM-DD')"
    strSql = strSql & "          And    a.Ptno      = b.Ptno(+)"
    strSql = strSql & "          And    c.JeobsuDt >= TO_DATE('" & sFrDate & "','YYYY-MM-DD')"
    strSql = strSql & "          And    c.JeobsuDt <= TO_DATE('" & sToDate & "','YYYY-MM-DD')"
    strSql = strSql & "          And    a.Ptno      = c.Ptno(+)"
    strSql = strSql & "          And    a.Slipno1   = c.Slipno1(+)"
    strSql = strSql & "          And    a.Slipno2   = c.Slipno2(+)"
    strSql = strSql & "          Group  By a.JEOBSUDT, a.PTNO, b.sname, b.Sex, b.AgeYY,  c.DeptCode, c.RoomCode,"
    strSql = strSql & "                    a.ITEMCD, a.RESULT1) "
    strSql = strSql & "  GROUP BY JeobsuDT, Ptno, Sname, Sex, ageyy, DeptCode, RoomCode"
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        sprChemi.Row = sprChemi.DataRowCnt + 1
        sprChemi.Col = 1: sprChemi.Text = adoSet.Fields("JeobsuDt").Value & ""
        sprChemi.Col = 2: sprChemi.Text = adoSet.Fields("DeptCode").Value & "/" & _
                                          adoSet.Fields("RoomCode").Value & ""
        sprChemi.Col = 3: sprChemi.Text = adoSet.Fields("Sname").Value & ""
        sprChemi.Col = 4: sprChemi.Text = adoSet.Fields("Sex").Value & "/" & _
                                          adoSet.Fields("AgeYY").Value & ""
        sprChemi.Col = 5:  sprChemi.Text = adoSet.Fields("Ca").Value & ""
        sprChemi.Col = 6:  sprChemi.Text = adoSet.Fields("P").Value & ""
        sprChemi.Col = 7:  sprChemi.Text = adoSet.Fields("Glu").Value & ""
        sprChemi.Col = 8:  sprChemi.Text = adoSet.Fields("BUN").Value & ""
        sprChemi.Col = 9:  sprChemi.Text = adoSet.Fields("Crea").Value & ""
        sprChemi.Col = 10: sprChemi.Text = adoSet.Fields("UA").Value & ""
        sprChemi.Col = 11: sprChemi.Text = adoSet.Fields("CholT").Value & ""
        sprChemi.Col = 12: sprChemi.Text = adoSet.Fields("ProT").Value & ""
        sprChemi.Col = 13: sprChemi.Text = adoSet.Fields("Alb").Value & ""
        sprChemi.Col = 14: sprChemi.Text = adoSet.Fields("AkP").Value & ""
        sprChemi.Col = 15: sprChemi.Text = adoSet.Fields("GOT").Value & ""
        sprChemi.Col = 16: sprChemi.Text = adoSet.Fields("GPT").Value & ""
        sprChemi.Col = 17: sprChemi.Text = adoSet.Fields("BiLT").Value & ""
        sprChemi.Col = 18: sprChemi.Text = adoSet.Fields("BiLD").Value & ""
        sprChemi.Col = 19: sprChemi.Text = adoSet.Fields("rGT").Value & ""
        sprChemi.Col = 20: sprChemi.Text = adoSet.Fields("TTT").Value & ""
        sprChemi.Col = 21: sprChemi.Text = adoSet.Fields("CPK").Value & ""
        sprChemi.Col = 22: sprChemi.Text = adoSet.Fields("LDH").Value & ""
        sprChemi.Col = 23: sprChemi.Text = adoSet.Fields("NH3").Value & ""
        sprChemi.Col = 24: sprChemi.Text = adoSet.Fields("Na").Value & ""
        sprChemi.Col = 25: sprChemi.Text = adoSet.Fields("K").Value & ""
        sprChemi.Col = 26: sprChemi.Text = adoSet.Fields("Cl").Value & ""
        sprChemi.Col = 27: sprChemi.Text = adoSet.Fields("Co2").Value & ""
        sprChemi.Col = 28: sprChemi.Text = adoSet.Fields("Amy").Value & ""
        sprChemi.Col = 29: sprChemi.Text = adoSet.Fields("Lip").Value & ""
        sprChemi.Col = 30: sprChemi.Text = adoSet.Fields("AcidPho").Value & ""
        sprChemi.Col = 31: sprChemi.Text = adoSet.Fields("Cholin").Value & ""
        sprChemi.Col = 32: sprChemi.Text = adoSet.Fields("IronS").Value & ""
        sprChemi.Col = 33: sprChemi.Text = adoSet.Fields("UIBC").Value & ""
        
        sprChemi.Col = 34: sprChemi.Text = adoSet.Fields("OsmoS").Value & ""
        sprChemi.Col = 35: sprChemi.Text = adoSet.Fields("OsmoU").Value & ""
        sprChemi.Col = 36: sprChemi.Text = adoSet.Fields("Mg").Value & ""
        sprChemi.Col = 37: sprChemi.Text = adoSet.Fields("Tg").Value & ""
        sprChemi.Col = 38: sprChemi.Text = adoSet.Fields("HDL").Value & ""
        sprChemi.Col = 39: sprChemi.Text = adoSet.Fields("LDL").Value & ""
        sprChemi.Col = 40: sprChemi.Text = adoSet.Fields("bLipo").Value & ""
        sprChemi.Col = 41: sprChemi.Text = adoSet.Fields("pLipid").Value & ""
        sprChemi.Col = 42: sprChemi.Text = adoSet.Fields("TLipid").Value & ""
        sprChemi.Col = 43: sprChemi.Text = adoSet.Fields("ADA").Value & ""
        sprChemi.Col = 44: sprChemi.Text = adoSet.Fields("Gluef").Value & ""
        sprChemi.Col = 45: sprChemi.Text = adoSet.Fields("Gluef2h").Value & ""
        sprChemi.Col = 46: sprChemi.Text = adoSet.Fields("OGttf").Value & ""
        
        sprChemi.Col = 47: sprChemi.Text = adoSet.Fields("OGttf").Value & ""
        sprChemi.Col = 48: sprChemi.Text = adoSet.Fields("OGtt3").Value & ""
        sprChemi.Col = 49: sprChemi.Text = adoSet.Fields("OGtt6").Value & ""
        sprChemi.Col = 50: sprChemi.Text = adoSet.Fields("OGtt9").Value & ""
        sprChemi.Col = 51: sprChemi.Text = adoSet.Fields("OGtt12").Value & ""
        sprChemi.Col = 52: sprChemi.Text = adoSet.Fields("OGtt18").Value & ""
        sprChemi.Col = 53: sprChemi.Text = adoSet.Fields("HbA1c").Value & ""
        sprChemi.Col = 54: sprChemi.Text = adoSet.Fields("Fruc").Value & ""
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    

End Sub

Private Sub Form_Load()

    GoSub SLip_Select
    dtFdate.Value = Dual_Date_Get("yyyy-MM-dd")
    dtTdate.Value = Dual_Date_Get("yyyy-MM-dd")
    
    Exit Sub
    
    

SLip_Select:
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_Specode"
    strSql = strSql & " WHERE  Codegu = '12'"
    strSql = strSql & " AND    Codeky < '90'"
'C    strSql = strSql & " AND    Codeky < '52'"
    strSql = strSql & " ORDER  BY Codeky"
    
    cmbSLip.Clear
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        cmbSLip.AddItem Trim(adoSet.Fields("Codeky").Value & "") & ". " & _
                        Trim(adoSet.Fields("Codenm").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
        
    Return

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1: Unload Me
    End Select
End Sub
