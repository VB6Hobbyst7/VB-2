VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmErrorCode 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Error Code ¼³¸í"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin Threed.SSPanel SSPanel2 
      Height          =   5205
      Left            =   6180
      TabIndex        =   2
      Top             =   660
      Width           =   3825
      _Version        =   65536
      _ExtentX        =   6747
      _ExtentY        =   9181
      _StockProps     =   15
      BackColor       =   15591915
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Begin VB.CommandButton cmdClose 
         Caption         =   "´Ý±â"
         Height          =   525
         Left            =   2820
         TabIndex        =   12
         Top             =   2280
         Width           =   825
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   525
         Left            =   1970
         TabIndex        =   11
         Top             =   2280
         Width           =   825
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "»èÁ¦"
         Height          =   525
         Left            =   1120
         TabIndex        =   10
         Top             =   2280
         Width           =   825
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "È®ÀÎ"
         Height          =   525
         Left            =   270
         TabIndex        =   9
         Top             =   2280
         Width           =   825
      End
      Begin VB.TextBox txtErrDesc 
         Appearance      =   0  'Æò¸é
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1710
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtChgErr 
         Appearance      =   0  'Æò¸é
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1710
         TabIndex        =   6
         Top             =   660
         Width           =   1935
      End
      Begin VB.TextBox txtTLAErr 
         Appearance      =   0  'Æò¸é
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1710
         TabIndex        =   4
         Top             =   255
         Width           =   1935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "¼³¸í"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   7
         Top             =   1125
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "º¯È¯ErrorCode"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   5
         Top             =   705
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "TLA ErrorCode"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   3
         Top             =   300
         Width           =   1365
      End
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   5205
      Left            =   30
      TabIndex        =   1
      Top             =   660
      Width           =   6105
      _Version        =   393216
      _ExtentX        =   10769
      _ExtentY        =   9181
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²Ã¼"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   3
      MaxRows         =   20
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "frmErrorCode.frx":0000
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   645
      Left            =   -390
      TabIndex        =   0
      Top             =   0
      Width           =   10410
      _Version        =   65536
      _ExtentX        =   18362
      _ExtentY        =   1138
      _StockProps     =   15
      Caption         =   "     Error Code"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²Ã¼"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Alignment       =   1
   End
End
Attribute VB_Name = "frmErrorCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Display_List()
    SQL = "Select tlaerror, chgerror, errordesc  from errorcode"
    res = db_select_Vas(gLocal, SQL, vasList)
    
    vasList.MaxRows = vasList.DataRowCnt
    
End Sub

Private Sub cmdClear_Click()
    txtTLAErr = ""
    txtChgErr = ""
    txtErrDesc = ""
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If Trim(txtTLAErr) = "" Then
        txtTLAErr.SetFocus
        Exit Sub
    End If
    
    
    SQL = "delete from errorcode  " & vbCrLf & _
          "where tlaerror = '" & Trim(txtTLAErr) & "'   "
    res = SendQuery(gLocal, SQL)
    
    Display_List
    cmdClear_Click

End Sub

Private Sub cmdSave_Click()
    If Trim(txtTLAErr) = "" Then
        txtTLAErr.SetFocus
        Exit Sub
    End If
    
    If Trim(txtChgErr) = "" Then
        txtChgErr.SetFocus
        Exit Sub
    End If
    

    
    SQL = "Select tlaerror  from errorcode  " & vbCrLf & _
          "where tlaerror = '" & Trim(txtTLAErr) & "'  "
    res = db_select_Col(gLocal, SQL)
    If Trim(gReadBuf(0)) = Trim(txtTLAErr) Then
        SQL = "update errorcode set chgerror = '" & Trim(txtChgErr) & "', errordesc = '" & Trim(txtErrDesc) & "' " & vbCrLf & _
              "where tlaerror = '" & Trim(txtTLAErr) & "'   "
        res = SendQuery(gLocal, SQL)
    Else
        SQL = "Insert into errorcode (tlaerror, chgerror, errordesc ) " & vbCrLf
        SQL = SQL & " Values ('" & Trim(txtTLAErr) & "', '" & Trim(txtChgErr) & "', '" & Trim(txtErrDesc) & "' ) "
        res = SendQuery(gLocal, SQL)
    End If
    
    Display_List
    cmdClear_Click
    
End Sub

Private Sub Form_Load()
    Display_List
    cmdClear_Click
End Sub

Private Sub txtErrDesc_GotFocus()
    SelectFocus txtErrDesc
End Sub

Private Sub txtErrDesc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdSave.SetFocus
    End If
End Sub

Private Sub txtTLAErr_GotFocus()
    SelectFocus txtTLAErr
End Sub

Private Sub txtTLAErr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtTLAErr) = "" Then
            txtTLAErr.SetFocus
        Else
            txtChgErr.SetFocus
        End If
    End If
    
End Sub

Private Sub txtChgErr_GotFocus()
    SelectFocus txtChgErr
End Sub

Private Sub txtChgErr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtChgErr) = "" Then
            txtChgErr.SetFocus
        Else
            txtErrDesc.SetFocus
        End If
    End If
End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then
        vasSort vasList, Col
        Exit Sub
    End If
        
    txtTLAErr = Trim(GetText(vasList, Row, 1))
    txtChgErr = Trim(GetText(vasList, Row, 2))
    txtErrDesc = Trim(GetText(vasList, Row, 3))
    
End Sub
