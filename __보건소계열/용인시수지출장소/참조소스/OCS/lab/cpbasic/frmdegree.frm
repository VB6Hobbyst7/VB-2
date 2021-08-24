VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmDegree 
   Caption         =   "정도관리 Item "
   ClientHeight    =   7485
   ClientLeft      =   135
   ClientTop       =   1170
   ClientWidth     =   11445
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
   ScaleHeight     =   7485
   ScaleWidth      =   11445
   WindowState     =   2  '최대화
   Begin FPSpreadADO.fpSpread vaSpread1 
      Height          =   5835
      Left            =   5490
      TabIndex        =   4
      Top             =   630
      Width           =   5055
      _Version        =   196608
      _ExtentX        =   8916
      _ExtentY        =   10292
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
      ScrollBars      =   2
      SpreadDesigner  =   "frmDegree.frx":0000
      Appearance      =   1
   End
   Begin FPSpreadADO.fpSpread ssDegree 
      Height          =   5505
      Left            =   135
      TabIndex        =   3
      Top             =   765
      Width           =   5235
      _Version        =   196608
      _ExtentX        =   9234
      _ExtentY        =   9710
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
      MaxCols         =   501
      ScrollBarExtMode=   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "frmDegree.frx":501A
      UserResize      =   1
      VisibleCols     =   500
      VisibleRows     =   500
      Appearance      =   1
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   315
      Left            =   3840
      TabIndex        =   2
      Top             =   360
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   556
      _StockProps     =   78
      Caption         =   "SSCommand1"
   End
   Begin VB.ComboBox cmbSlip 
      Height          =   300
      Left            =   1380
      Style           =   2  '드롭다운 목록
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Slip 종류"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   420
      Width           =   975
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "Quit"
   End
End
Attribute VB_Name = "frmDegree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbSlip_Click()
    Dim sCode       As String
    
    If cmbSlip.ListIndex = -1 Then Exit Sub
    
    sCode = Left(cmbSlip.Text, 2)
    
    StrSql = ""
    StrSql = StrSql & " SELECT a.RowID, a.*"
    StrSql = StrSql & " FROM   TWEXAM_Degree_Item a"
    StrSql = StrSql & " WHERE  a.ItemCd  LIKE '" & sCode & "%'"
    StrSql = StrSql & " ORDER  BY a.GeomsaCode"
    
    ssDegree.MaxRows = 0
    If False = adoSetOpen(StrSql, adoSet) Then Exit Sub
    ssDegree.MaxRows = adoSet.RecordCount
    
    Do Until adoSet.EOF
        ssDegree.Row = ssDegree.DataRowCnt + 1
        ssDegree.Col = 1:  ssDegree.Text = adoSet.Fields("RowID").Value & ""
        ssDegree.Col = 2:  ssDegree.Text = adoSet.Fields("GeomsaCode").Value & ""
        ssDegree.Col = 3:  ssDegree.Text = adoSet.Fields("GeomsaName").Value & ""
        ssDegree.Col = 4:  ssDegree.Text = adoSet.Fields("ItemCD").Value & ""
        ssDegree.Col = 5:  ssDegree.Text = adoSet.Fields("ItemNm").Value & ""
        ssDegree.Col = 6:  ssDegree.Text = adoSet.Fields("Danwi").Value & ""
        ssDegree.Col = 7:  ssDegree.Text = adoSet.Fields("MinRef").Value & ""
        ssDegree.Col = 8:  ssDegree.Text = adoSet.Fields("MaxRef").Value & ""
        ssDegree.Col = 9:  ssDegree.Text = adoSet.Fields("GbLevel").Value & ""
        ssDegree.Col = 10: ssDegree.Text = adoSet.Fields("GeomJangbi").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
End Sub

Private Sub Form_Load()
    
    GoSub Get_Specode12
    Exit Sub
    
'/_______________________________________________
Get_Specode12:
    StrSql = ""
    StrSql = StrSql & " SELECT *"
    StrSql = StrSql & " FROM   TWEXAM_SPECODE"
    StrSql = StrSql & " WHERE  Codegu = '12'"
    StrSql = StrSql & " ORDER  BY Codeky"
    
    If False = adoSetOpen(StrSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        cmbSlip.AddItem Trim(adoSet.Fields("Codeky").Value & "") & ". " & _
                        Trim(adoSet.Fields("Codenm").Value & "")
        adoSet.MoveNext
    Loop
    
    Call adoSetClose(adoSet)
    
    Return
    

End Sub

Private Sub mnuQuit_Click()
    Unload Me
    
End Sub

