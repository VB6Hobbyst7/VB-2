VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmReport 
   Caption         =   "검사종목별 결과보고서"
   ClientHeight    =   6075
   ClientLeft      =   2100
   ClientTop       =   1785
   ClientWidth     =   9330
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
   ScaleHeight     =   6075
   ScaleWidth      =   9330
   WindowState     =   2  '최대화
   Begin FPSpreadADO.fpSpread sprReport 
      Height          =   5055
      Left            =   315
      TabIndex        =   8
      Top             =   1755
      Width           =   10770
      _Version        =   196608
      _ExtentX        =   18997
      _ExtentY        =   8916
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
      SpreadDesigner  =   "frmReport.frx":0000
      Appearance      =   1
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1410
      Left            =   315
      TabIndex        =   0
      Top             =   180
      Width           =   8610
      _Version        =   65536
      _ExtentX        =   15187
      _ExtentY        =   2487
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
      Begin VB.TextBox txtSname 
         Height          =   285
         Left            =   2700
         TabIndex        =   6
         Top             =   630
         Width           =   1230
      End
      Begin VB.ComboBox cmbSLip 
         Height          =   300
         Left            =   4455
         TabIndex        =   5
         Top             =   225
         Width           =   2490
      End
      Begin MSComCtl2.DTPicker dtToDate 
         Height          =   330
         Left            =   2745
         TabIndex        =   4
         Top             =   225
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24576003
         CurrentDate     =   36442
      End
      Begin MSComCtl2.DTPicker dtFrDate 
         Height          =   330
         Left            =   1395
         TabIndex        =   3
         Top             =   225
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24576003
         CurrentDate     =   36442
      End
      Begin VB.TextBox txtPtno 
         Height          =   285
         Left            =   1395
         TabIndex        =   2
         Top             =   630
         Width           =   1230
      End
      Begin VB.Label Label2 
         Caption         =   "기준일자:"
         Height          =   240
         Left            =   450
         TabIndex        =   7
         Top             =   270
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "등록번호:"
         Height          =   195
         Left            =   450
         TabIndex        =   1
         Top             =   675
         Width           =   870
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    
    GoSub Date_Settting
    GoSub SLip_Setting
    Exit Sub
    
    
Date_Settting:
    dtFrDate.Value = Dual_Date_Get("yyyy-MM-dd")
    dtToDate.Value = Dual_Date_Get("yyyy-MM-dd")
    
    Return
    
SLip_Setting:
    StrSql = ""
    StrSql = StrSql & " SELECT *"
    StrSql = StrSql & " FROM   TWEXAM_Specode"
    StrSql = StrSql & " WHERE  Codegu = '12'"
    StrSql = StrSql & " AND    Codeky < '52'"
    StrSql = StrSql & " Order  By Codeky"
    If False = adoSetOpen(StrSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        cmbSLip.AddItem Trim(adoSet.Fields("Codeky").Value & "") & " " & _
                             adoSet.Fields("Codenm").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    
    Return

End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

Private Sub txtPtno_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Trim(txtPtno.Text) = "" Then Exit Sub
        txtPtno.Text = Format(txtPtno.Text, "00000000")
        GoSub Check_Patient_Data
    End If
    Exit Sub
    
    
Check_Patient_Data:
'o  StrSql = ""
'o  StrSql = StrSql & " SELECT /*+ INDEX (TWBas_Patient INDEX_PATIENT0) */"

    StrSql = ""
    StrSql = StrSql & " SELECT Sname, Sex, Jumin1, Jumin2, LastDate"
    StrSql = StrSql & " FROM   TWBas_Patient"
    StrSql = StrSql & " WHERE  Ptno = '" & txtPtno.Text & "'"
    If False = adoSetOpen(StrSql, adoSet) Then Return
    
    txtSname.Text = adoSet.Fields("Sname").Value & ""
    Call adoSetClose(adoSet)
    
    Return
    
    
End Sub
