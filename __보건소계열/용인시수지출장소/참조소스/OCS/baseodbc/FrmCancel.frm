VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmCancel 
   Caption         =   "내시경검사 취소확인"
   ClientHeight    =   735
   ClientLeft      =   4560
   ClientTop       =   1950
   ClientWidth     =   3090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   3090
   Begin Threed.SSCommand CmdTcancel 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "검사취소"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
   Begin Threed.SSCommand CmdExit 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "종  료"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
End
Attribute VB_Name = "FrmCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim rs          As ADODB.Recordset
Dim Rs1         As ADODB.Recordset

Private Sub CmdExit_Click()
    
    Unload Me
    FrmReserved.Cmdprint.Enabled = True
    
End Sub

Private Sub CmdTcancel_Click()

    Dim strPtno                      As String
    Dim strName                      As String
    Dim strSex                       As String
    Dim strPDate                     As String
    Dim StrDate                      As String
    Dim strCode                      As String
    Dim i                            As Integer

    FrmReserved.Cmdprint.Enabled = False
    
    Unload Me

    FrmReserved.SS1.Row = FrmReserved.SS1.ActiveRow
    FrmReserved.SS1.Col = 1:
    strPtno = Trim(FrmReserved.SS1.Text)
    FrmReserved.SS1.Col = 5:
    StrDate = Trim(FrmReserved.SS1.Text)
    FrmReserved.SS1.Col = 14:
    strCode = Trim(FrmReserved.SS1.Text)
    
    
    strSQL = ""
    strSQL = "Select To_Char(SysDate,'YYYY-MM-DD') SDate,  "
    strSQL = strSQL & " To_Char(SysDate, 'HH24:MM') STime From Dual"
    Result = adoSQL(strSQL)

    strPDate = Date_Format(AdoGetString(rs, "SDate", 0))
    
    strSQL = ""
    strSQL = " SELECT A.GBJOB, A.PTNO, P.SNAME, A.ORDERCODE, B.ORDERNAMES, B.SUCODE, D.DRNAME, "
    strSQL = strSQL & " A.RDATE,  TO_CHAR(VDATE, 'HH24:MI') RTIME "
    strSQL = strSQL & " FROM TWENDO_JUPMST A, TWOCS_ORDERCODE B, "
    strSQL = strSQL & "  TW_MIS_PMPA.TWBAS_DOCTOR  D, TW_MIS_PMPA.TWBAS_PATIENT P "
    strSQL = strSQL & " WHERE A.OrderCode = B.OrderCode   "
    strSQL = strSQL & " AND A.PTNO = '" & strPtno & "'"
    strSQL = strSQL & " AND A.PTNO = P.PTNO "
    strSQL = strSQL & " AND A.JDATE = TO_DATE('" & StrDate & "','YYYY-MM-DD') "
    strSQL = strSQL & " AND B.Slipno    = '0040' "               ' 0044..' =>한림병원 0040
    strSQL = strSQL & " AND A.OrderCode = '" & strCode & "'"
    strSQL = strSQL & " AND A.DrCode    = D.DrCode(+)   "
    strSQL = strSQL & " AND A.GBSUNAP = '3'"
'    strSql = strSql & " AND A.ResultDate IS NULL   "
    Result = AdoOpenSet(Rs1, strSQL)

    If rowindicator = 0 Then
        Exit Sub
    End If

    Res_Print.SS3.Col = 4
    Res_Print.SS3.Row = 7:  Res_Print.SS3.Text = AdoGetString(Rs1, "SNAME", 0)
    Res_Print.SS3.Row = 8:  Res_Print.SS3.Text = AdoGetString(Rs1, "PTNO", 0)
    Res_Print.SS3.Row = 11: Res_Print.SS3.Text = AdoGetString(Rs1, "SUCODE", 0) & AdoGetString(Rs1, "ORDERNAMES", 0)
    Res_Print.SS3.Row = 12: Res_Print.SS3.Text = AdoGetString(Rs1, "DRNAME", 0)
    Res_Print.SS3.Row = 13:
    Res_Print.SS3.Text = Format(AdoGetString(Rs1, "RDATE", 0), "YYYY-MM-DD") & " " & Trim(AdoGetString(Rs1, "RTIME", 0)) 'Format(AdoGetString(Rs1, "RTIME", 0), "HH:MM AMPM")
        
    Res_Print.Left = 0
    Res_Print.Top = 0
    Res_Print.Show 1
    
    AdoCloseSet rs
    AdoCloseSet Rs1

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Unload Me
    
End Sub


