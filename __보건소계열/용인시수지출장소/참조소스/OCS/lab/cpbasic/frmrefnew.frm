VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRefNew 
   Appearance      =   0  '평면
   BackColor       =   &H80000005&
   Caption         =   "신규추가화면"
   ClientHeight    =   4935
   ClientLeft      =   1575
   ClientTop       =   2385
   ClientWidth     =   3420
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
   ScaleHeight     =   4935
   ScaleWidth      =   3420
   Begin VB.TextBox txtItemName 
      BackColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   900
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "Text2"
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox txtItemCode 
      BackColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   120
      Width           =   840
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   4035
      Left            =   60
      TabIndex        =   10
      Top             =   540
      Width           =   3255
      _Version        =   65536
      _ExtentX        =   5741
      _ExtentY        =   7117
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
      Begin VB.TextBox txtRowID 
         Appearance      =   0  '평면
         BackColor       =   &H80000000&
         Height          =   270
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   120
         Width           =   2835
      End
      Begin MSComCtl2.DTPicker dtAppdate 
         Height          =   315
         Left            =   1560
         TabIndex        =   11
         Top             =   420
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24510467
         CurrentDate     =   36325
      End
      Begin VB.TextBox txtFmax 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txtFmin 
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   2340
         Width           =   1095
      End
      Begin VB.TextBox txtMmax 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   1980
         Width           =   1095
      End
      Begin VB.TextBox txtMmin 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtAgemax 
         Height          =   315
         Left            =   2280
         TabIndex        =   2
         Top             =   1260
         Width           =   675
      End
      Begin VB.TextBox txtAgemin 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   1260
         Width           =   675
      End
      Begin VB.TextBox txtAppGubun 
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         Top             =   840
         Width           =   495
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   915
         Left            =   2340
         TabIndex        =   12
         Top             =   3060
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   1614
         _StockProps     =   78
         Caption         =   "종료"
         Picture         =   "frmRefNew.frx":0000
      End
      Begin Threed.SSCommand cmdInsert 
         Height          =   915
         Left            =   600
         TabIndex        =   7
         Top             =   3060
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   1614
         _StockProps     =   78
         Caption         =   "입력확인"
         Picture         =   "frmRefNew.frx":08DA
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   915
         Left            =   1500
         TabIndex        =   18
         Top             =   3060
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   1614
         _StockProps     =   78
         Caption         =   "삭제확인"
      End
      Begin VB.Label Label7 
         Caption         =   "여)최저/최고"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   2400
         Width           =   1155
      End
      Begin VB.Label Label5 
         Caption         =   "남)최저/최고"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   1740
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "연령(min/max)"
         Height          =   255
         Left            =   180
         TabIndex        =   15
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "적용구분"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "적용일자"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   540
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmRefNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
    
End Sub

Private Sub cmdInsert_Click()
    Dim sAppdate        As String
    
    sAppdate = Format(dtAppdate.Value, "yyyy-MM-dd")
    
    If Val(txtAgemin.Text) = 0 Or Val(txtAgemax.Text) = 0 Then Exit Sub
    
    If Trim(txtRowID.Text) = "" Then
        GoSub Ref_Insert
    Else
        GoSub Ref_Update
    End If
    Exit Sub
    
Ref_Insert:
    strSql = ""
    strSql = strSql & " INSERT INTO TWEXAM_RefData"
    strSql = strSql & "        (iTemCode, appDate, appGubun, ageMin, ageMax, "
    strSql = strSql & "         M_min,    M_max,   F_min,    F_max)"
    strSql = strSql & " VALUES ('" & txtItemCode.Text & "',"
    strSql = strSql & "              TO_DATE('" & sAppdate & "','YYYY-MM-DD'),"
    strSql = strSql & "         '" & txtAppGubun & "',"
    strSql = strSql & "          " & Val(txtAgemin.Text) & ","
    strSql = strSql & "          " & Val(txtAgemax.Text) & ","
    strSql = strSql & "         '" & Trim(txtMmin.Text) & "',"
    strSql = strSql & "         '" & Trim(txtMmax.Text) & "',"
    strSql = strSql & "         '" & Trim(txtFmin.Text) & "',"
    strSql = strSql & "         '" & Trim(txtFmax.Text) & "')"
    
    
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
        GoSub Form_Clear_Routine
    Else
        MsgBox "해당 Data 가 입력되지 않았습니다!......."
        adoConnect.RollbackTrans
    End If
    Return

Ref_Update:
    strSql = ""
    strSql = strSql & " UPDATE TWEXAM_RefData"
    strSql = strSql & " SET    appDate  = TO_DATE('" & sAppdate & "','yyyy-MM-dd'),"
    strSql = strSql & "        appGubun = '" & txtAppGubun.Text & "',"
    strSql = strSql & "        ageMin   =  " & Val(txtAgemin.Text) & ","
    strSql = strSql & "        ageMax   =  " & Val(txtAgemax.Text) & ","
    strSql = strSql & "        M_min    =  " & Val(txtMmin.Text) & ","
    strSql = strSql & "        M_max    =  " & Val(txtMmax.Text) & ","
    strSql = strSql & "        F_min    =  " & Val(txtFmin.Text) & ","
    strSql = strSql & "        F_max    =  " & Val(txtFmax.Text)
    strSql = strSql & " WHERE  RowID    = '" & txtRowID.Text & "'"
    
    
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
        GoSub Form_Clear_Routine
    Else
        MsgBox "해당 Data 가 수정되지 않았습니다!......."
        adoConnect.RollbackTrans
    End If
    Return
    
    
Form_Clear_Routine:
    txtAppGubun.Text = ""
    txtAgemin.Text = ""
    txtAgemax.Text = ""
    txtMmin.Text = ""
    txtMmax.Text = ""
    txtFmin.Text = ""
    txtFmax.Text = ""
    
    Return
    
End Sub

Private Sub Form_Load()
    
    txtItemCode.Text = frmNormal1.txtCode.Text
    txtItemName.Text = frmNormal1.txtName.Text
    dtAppdate.Value = Dual_Date_Get("YYYY-MM-DD")
    
    If hWndReturn = frmNormal1.tvRef.hwnd Then
        txtRowID.Text = ""
    ElseIf hWndReturn = frmNormal1.ssRefData.hwnd Then
        frmNormal1.ssRefData.Row = frmNormal1.ssRefData.ActiveRow
        
        frmNormal1.ssRefData.Col = 1: txtRowID.Text = frmNormal1.ssRefData.Text
        frmNormal1.ssRefData.Col = 2: dtAppdate.Value = Format(frmNormal1.ssRefData.Text, "yyyy-MM-dd")
        frmNormal1.ssRefData.Col = 4: txtAppGubun.Text = frmNormal1.ssRefData.Text
        frmNormal1.ssRefData.Col = 5: txtAgemin.Text = frmNormal1.ssRefData.Text
        frmNormal1.ssRefData.Col = 6: txtAgemax.Text = frmNormal1.ssRefData.Text
        frmNormal1.ssRefData.Col = 7: txtMmin.Text = frmNormal1.ssRefData.Text
        frmNormal1.ssRefData.Col = 8: txtMmax.Text = frmNormal1.ssRefData.Text
        frmNormal1.ssRefData.Col = 9: txtFmin.Text = frmNormal1.ssRefData.Text
        frmNormal1.ssRefData.Col = 10: txtFmax.Text = frmNormal1.ssRefData.Text
    End If
    
End Sub

Private Sub mnuQuit_Click()
    Unload Me
    
End Sub
