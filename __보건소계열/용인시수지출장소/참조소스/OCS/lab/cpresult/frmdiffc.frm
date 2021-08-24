VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{B16553C3-06DB-101B-85B2-0000C009BE81}#1.0#0"; "SPIN32.OCX"
Begin VB.Form frmDiffc 
   Caption         =   "Differential CountForm"
   ClientHeight    =   4245
   ClientLeft      =   5640
   ClientTop       =   2445
   ClientWidth     =   6105
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6105
   Begin VB.Frame Frame1 
      Caption         =   "Max. Count"
      Height          =   555
      Left            =   2970
      TabIndex        =   43
      Top             =   90
      Width           =   2940
      Begin VB.OptionButton opt100 
         Caption         =   "100"
         Height          =   240
         Left            =   1305
         TabIndex        =   0
         Top             =   225
         Value           =   -1  'True
         Width           =   690
      End
      Begin VB.OptionButton opt50 
         Caption         =   "50"
         Height          =   240
         Left            =   540
         TabIndex        =   1
         Top             =   225
         Width           =   690
      End
      Begin VB.OptionButton opt500 
         Caption         =   "500"
         Height          =   240
         Left            =   2070
         TabIndex        =   2
         Top             =   225
         Width           =   645
      End
   End
   Begin Threed.SSPanel panelDiff 
      Height          =   3435
      Left            =   180
      TabIndex        =   3
      Top             =   675
      Width           =   5775
      _Version        =   65536
      _ExtentX        =   10186
      _ExtentY        =   6059
      _StockProps     =   15
      Caption         =   " False"
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
      Alignment       =   2
      Begin Spin.SpinButton spinPer 
         Height          =   270
         Index           =   0
         Left            =   675
         TabIndex        =   58
         Top             =   705
         Width           =   195
         _Version        =   65536
         _ExtentX        =   344
         _ExtentY        =   476
         _StockProps     =   73
      End
      Begin VB.TextBox txtTitle 
         BackColor       =   &H00404000&
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   0
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Tag             =   "11012201"
         Text            =   "Atyp.Lymph(7)"
         Top             =   135
         Width           =   1320
      End
      Begin VB.TextBox txtTitle 
         BackColor       =   &H00404000&
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   1
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Tag             =   "11012202"
         Text            =   "Blast(8)"
         Top             =   135
         Width           =   1320
      End
      Begin VB.TextBox txtTitle 
         BackColor       =   &H00404000&
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   2
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Tag             =   "11012203"
         Text            =   "Imm.Cell(9)"
         Top             =   135
         Width           =   1320
      End
      Begin VB.TextBox txtTitle 
         BackColor       =   &H00404000&
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   3
         Left            =   4230
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Tag             =   "11012204"
         Text            =   "Promyelocyt(-)"
         Top             =   135
         Width           =   1320
      End
      Begin VB.TextBox txtTitle 
         BackColor       =   &H00404000&
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   4
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   "Myelocyte(4)"
         Top             =   1080
         Width           =   1320
      End
      Begin VB.TextBox txtTitle 
         BackColor       =   &H00404000&
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   5
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Text            =   "Metamyelo(5)"
         Top             =   1080
         Width           =   1320
      End
      Begin VB.TextBox txtTitle 
         BackColor       =   &H00404000&
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   6
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "Band N(6)"
         Top             =   1080
         Width           =   1320
      End
      Begin VB.TextBox txtTitle 
         BackColor       =   &H00404000&
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   7
         Left            =   4230
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "Seg N(+)"
         Top             =   1080
         Width           =   1320
      End
      Begin VB.TextBox txtTitle 
         BackColor       =   &H00404000&
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   8
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "Lymphocyte(1)"
         Top             =   2070
         Width           =   1320
      End
      Begin VB.TextBox txtTitle 
         BackColor       =   &H00404000&
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   9
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "Monocyte(2)"
         Top             =   2070
         Width           =   1320
      End
      Begin VB.TextBox txtTitle 
         BackColor       =   &H00404000&
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   10
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "Eosinophil(3)"
         Top             =   2070
         Width           =   1320
      End
      Begin VB.TextBox txtTitle 
         BackColor       =   &H00404000&
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   11
         Left            =   4230
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "Basophil(.)"
         Top             =   2070
         Width           =   1320
      End
      Begin VB.TextBox txtCount 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   29
         Top             =   420
         Width           =   960
      End
      Begin VB.TextBox txtCount 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1530
         TabIndex        =   28
         Top             =   420
         Width           =   960
      End
      Begin VB.TextBox txtCount 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2880
         TabIndex        =   27
         Top             =   420
         Width           =   960
      End
      Begin VB.TextBox txtCount 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   4230
         TabIndex        =   26
         Top             =   420
         Width           =   960
      End
      Begin VB.TextBox txtCount 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   180
         TabIndex        =   25
         Top             =   1365
         Width           =   960
      End
      Begin VB.TextBox txtCount 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   1530
         TabIndex        =   24
         Top             =   1365
         Width           =   960
      End
      Begin VB.TextBox txtCount 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2880
         TabIndex        =   23
         Top             =   1365
         Width           =   960
      End
      Begin VB.TextBox txtCount 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   4230
         TabIndex        =   22
         Top             =   1365
         Width           =   960
      End
      Begin VB.TextBox txtCount 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   180
         TabIndex        =   21
         Top             =   2355
         Width           =   960
      End
      Begin VB.TextBox txtCount 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   1530
         TabIndex        =   20
         Top             =   2355
         Width           =   960
      End
      Begin VB.TextBox txtCount 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   2880
         TabIndex        =   19
         Top             =   2355
         Width           =   960
      End
      Begin VB.TextBox txtCount 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   4230
         TabIndex        =   18
         Top             =   2355
         Width           =   960
      End
      Begin VB.TextBox txtPercent 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   180
         TabIndex        =   17
         Tag             =   "11012208"
         Top             =   705
         Width           =   510
      End
      Begin VB.TextBox txtPercent 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   1530
         TabIndex        =   16
         Tag             =   "11012209"
         Top             =   705
         Width           =   510
      End
      Begin VB.TextBox txtPercent 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   270
         Index           =   2
         Left            =   2880
         TabIndex        =   15
         Tag             =   "11012210"
         Top             =   705
         Width           =   510
      End
      Begin VB.TextBox txtPercent 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   270
         Index           =   3
         Left            =   4230
         TabIndex        =   14
         Tag             =   "11012211"
         Top             =   705
         Width           =   510
      End
      Begin VB.TextBox txtPercent 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   270
         Index           =   4
         Left            =   180
         TabIndex        =   13
         Tag             =   "11012212"
         Top             =   1650
         Width           =   510
      End
      Begin VB.TextBox txtPercent 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   270
         Index           =   5
         Left            =   1530
         TabIndex        =   12
         Tag             =   "11012213"
         Top             =   1650
         Width           =   510
      End
      Begin VB.TextBox txtPercent 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   270
         Index           =   6
         Left            =   2880
         TabIndex        =   11
         Tag             =   "11012201"
         Top             =   1650
         Width           =   510
      End
      Begin VB.TextBox txtPercent 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   270
         Index           =   7
         Left            =   4230
         TabIndex        =   10
         Tag             =   "11012202"
         Top             =   1650
         Width           =   510
      End
      Begin VB.TextBox txtPercent 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   270
         Index           =   8
         Left            =   180
         TabIndex        =   9
         Tag             =   "11012203"
         Top             =   2640
         Width           =   510
      End
      Begin VB.TextBox txtPercent 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   270
         Index           =   9
         Left            =   1530
         TabIndex        =   8
         Tag             =   "11012204"
         Top             =   2640
         Width           =   510
      End
      Begin VB.TextBox txtPercent 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   270
         Index           =   10
         Left            =   2880
         TabIndex        =   7
         Tag             =   "11012205"
         Top             =   2640
         Width           =   510
      End
      Begin VB.TextBox txtPercent 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   270
         Index           =   11
         Left            =   4230
         TabIndex        =   6
         Tag             =   "11012206"
         Top             =   2640
         Width           =   510
      End
      Begin VB.TextBox txtTotalCount 
         Alignment       =   1  '오른쪽 맞춤
         Enabled         =   0   'False
         Height          =   285
         Left            =   3375
         TabIndex        =   5
         Top             =   3015
         Width           =   780
      End
      Begin VB.TextBox txtTotalPercent 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4230
         TabIndex        =   4
         Top             =   3015
         Width           =   780
      End
      Begin Spin.SpinButton spinPer 
         Height          =   270
         Index           =   1
         Left            =   2025
         TabIndex        =   59
         Top             =   705
         Width           =   195
         _Version        =   65536
         _ExtentX        =   344
         _ExtentY        =   476
         _StockProps     =   73
      End
      Begin Spin.SpinButton spinPer 
         Height          =   270
         Index           =   2
         Left            =   3375
         TabIndex        =   60
         Top             =   705
         Width           =   195
         _Version        =   65536
         _ExtentX        =   344
         _ExtentY        =   476
         _StockProps     =   73
      End
      Begin Spin.SpinButton spinPer 
         Height          =   270
         Index           =   3
         Left            =   4725
         TabIndex        =   61
         Top             =   705
         Width           =   195
         _Version        =   65536
         _ExtentX        =   344
         _ExtentY        =   476
         _StockProps     =   73
      End
      Begin Spin.SpinButton spinPer 
         Height          =   270
         Index           =   4
         Left            =   675
         TabIndex        =   62
         Top             =   1650
         Width           =   195
         _Version        =   65536
         _ExtentX        =   344
         _ExtentY        =   476
         _StockProps     =   73
      End
      Begin Spin.SpinButton spinPer 
         Height          =   270
         Index           =   5
         Left            =   2025
         TabIndex        =   63
         Top             =   1650
         Width           =   195
         _Version        =   65536
         _ExtentX        =   344
         _ExtentY        =   476
         _StockProps     =   73
      End
      Begin Spin.SpinButton spinPer 
         Height          =   270
         Index           =   6
         Left            =   3375
         TabIndex        =   64
         Top             =   1650
         Width           =   195
         _Version        =   65536
         _ExtentX        =   344
         _ExtentY        =   476
         _StockProps     =   73
      End
      Begin Spin.SpinButton spinPer 
         Height          =   270
         Index           =   7
         Left            =   4725
         TabIndex        =   65
         Top             =   1650
         Width           =   195
         _Version        =   65536
         _ExtentX        =   344
         _ExtentY        =   476
         _StockProps     =   73
      End
      Begin Spin.SpinButton spinPer 
         Height          =   270
         Index           =   8
         Left            =   675
         TabIndex        =   66
         Top             =   2640
         Width           =   195
         _Version        =   65536
         _ExtentX        =   344
         _ExtentY        =   476
         _StockProps     =   73
      End
      Begin Spin.SpinButton spinPer 
         Height          =   270
         Index           =   9
         Left            =   2025
         TabIndex        =   67
         Top             =   2640
         Width           =   195
         _Version        =   65536
         _ExtentX        =   344
         _ExtentY        =   476
         _StockProps     =   73
      End
      Begin Spin.SpinButton spinPer 
         Height          =   270
         Index           =   10
         Left            =   3375
         TabIndex        =   68
         Top             =   2640
         Width           =   195
         _Version        =   65536
         _ExtentX        =   344
         _ExtentY        =   476
         _StockProps     =   73
      End
      Begin Spin.SpinButton spinPer 
         Height          =   270
         Index           =   11
         Left            =   4725
         TabIndex        =   69
         Top             =   2640
         Width           =   195
         _Version        =   65536
         _ExtentX        =   344
         _ExtentY        =   476
         _StockProps     =   73
      End
      Begin VB.Label Label2 
         Caption         =   "Total Count/Percentage"
         Height          =   195
         Left            =   1260
         TabIndex        =   57
         Top             =   3060
         Width           =   1995
      End
      Begin VB.Line Line1 
         X1              =   180
         X2              =   5490
         Y1              =   2970
         Y2              =   2970
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         Height          =   150
         Index           =   12
         Left            =   5040
         TabIndex        =   56
         Top             =   3060
         Width           =   105
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         Height          =   150
         Index           =   11
         Left            =   4950
         TabIndex        =   55
         Top             =   2700
         Width           =   105
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         Height          =   150
         Index           =   10
         Left            =   3600
         TabIndex        =   54
         Top             =   2700
         Width           =   105
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         Height          =   150
         Index           =   9
         Left            =   2250
         TabIndex        =   53
         Top             =   2700
         Width           =   105
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         Height          =   150
         Index           =   8
         Left            =   900
         TabIndex        =   52
         Top             =   2700
         Width           =   105
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         Height          =   150
         Index           =   7
         Left            =   4950
         TabIndex        =   51
         Top             =   1710
         Width           =   105
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         Height          =   150
         Index           =   6
         Left            =   3600
         TabIndex        =   50
         Top             =   1710
         Width           =   105
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         Height          =   150
         Index           =   5
         Left            =   2250
         TabIndex        =   49
         Top             =   1710
         Width           =   105
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         Height          =   150
         Index           =   4
         Left            =   900
         TabIndex        =   48
         Top             =   1710
         Width           =   105
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         Height          =   150
         Index           =   3
         Left            =   4995
         TabIndex        =   47
         Top             =   765
         Width           =   105
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         Height          =   150
         Index           =   2
         Left            =   3645
         TabIndex        =   46
         Top             =   765
         Width           =   105
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         Height          =   150
         Index           =   1
         Left            =   2295
         TabIndex        =   45
         Top             =   765
         Width           =   105
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         Height          =   150
         Index           =   0
         Left            =   900
         TabIndex        =   44
         Top             =   765
         Width           =   105
      End
   End
   Begin MSForms.CommandButton cmdMoveSet 
      Height          =   465
      Left            =   180
      TabIndex        =   42
      Top             =   180
      Width           =   1860
      Caption         =   "MoveSet"
      PicturePosition =   327683
      Size            =   "3281;820"
      Picture         =   "frmDiffc.frx":0000
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmDiffc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdMoveSet_Click()
    Dim j       As Integer
    
    If Val(txtTotalPercent.Text) <> 100 Then
        If vbNo = MsgBox("입력하신 Data 의 합계를 100%에 맞추십시오!." & vbCrLf & _
                             "무시하고 입력하시겠습니까?", _
                              vbYesNo + vbInformation, _
                             "100% 맞춤 확인 MessageBox") Then Exit Sub
    End If
    
    For i = 1 To frmResult.sprSLip.DataRowCnt
        For j = 0 To 11
            frmResult.sprSLip.Row = i
            frmResult.sprSLip.Col = 11
            If Trim(frmResult.sprSLip.Text) = Trim(txtPercent(j).Tag) Then
                frmResult.sprSLip.Col = 2: frmResult.sprSLip.Text = Trim(txtPercent(j).Text)
            End If
        Next
    Next

    Unload Me
    
End Sub

Private Sub Form_Activate()
    If opt50.Value = True Then opt50.SetFocus
    
    If opt100.Value = True Then opt100.SetFocus
    
    If opt500.Value = True Then opt500.SetFocus

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim iRow        As Integer
    Dim iCnt        As Integer
    Dim iCol        As Integer
    Dim iLimit      As Integer
    Dim iTotal      As Integer
    Dim j           As Integer
    Dim iCntIndex   As Integer
    Dim iPerIndex   As Integer
    
    
    iRow = 0
    
    If opt50.Value = True Then
        iLimit = 50: opt50.SetFocus: End If
    If opt100.Value = True Then
        iLimit = 100: opt100.SetFocus: End If
    If opt500.Value = True Then
        iLimit = 500: opt500.SetFocus: End If
    
    
    Select Case KeyCode
        Case 103:  iCntIndex = 0      'Atyp.Lymph     7
        Case 104:  iCntIndex = 1      'Blast          8
        Case 105:  iCntIndex = 2      'Imm.Cell       9
        Case 109:  iCntIndex = 3      'Promyelocyte   -
                   
        Case 100:  iCntIndex = 4      'Myelocyte      4
        Case 101:  iCntIndex = 5      'Metamyelo      5
        Case 102:  iCntIndex = 6      'Band N         6
        Case 107:  iCntIndex = 7      'Seg N          +
        
        Case 97:   iCntIndex = 8       'Lymphocyte     1
        Case 98:   iCntIndex = 9       'Monocyte       2
        Case 99:   iCntIndex = 10      'Eosinophil     3
        Case 110:  iCntIndex = 11      'Basophil       .
        Case Else:
                    If KeyCode = 13 Then
                        GoSub Percen_ReCal_Sub
                    End If
                    Exit Sub
    End Select
    
    If Val(txtTotalCount.Text) >= iLimit Then
        MsgBox "이미 설정하신 숫자에 접근하였습니다!.", vbCritical
        Exit Sub
    End If
    
    
    GoSub Count_Calc_Sub
    GoSub PerCent_Calc_Sub
    If opt50.Value = True Then opt50.SetFocus
    If opt100.Value = True Then opt100.SetFocus
    If opt500.Value = True Then opt500.SetFocus
    
    Exit Sub

'/-------------------------------------------------------------------
Count_Calc_Sub:
    'Edit TextBox Sum
    txtCount(iCntIndex).Text = Val(txtCount(iCntIndex).Text) + 1
    
    'Total Count Sum
    iTotal = 0
    For i = 0 To 11
        iTotal = iTotal + Val(txtCount(i).Text)
    Next
    txtTotalCount.Text = iTotal
    Return
    
    
PerCent_Calc_Sub:
    'Percent Textbox
    For i = 0 To 11
        If txtCount(i).Text <> "" Then
            txtPercent(i).Text = Val(txtCount(i).Text) * 100 / Val(txtTotalCount.Text)
            txtPercent(i).Text = Round(Val(txtPercent(i).Text))
        End If
    Next
    
    'Percent Sum
    iTotal = 0
    For i = 0 To 11
        iTotal = iTotal + Val(txtPercent(i).Text)
    Next
    txtTotalPercent.Text = iTotal
    Return
    
Percen_ReCal_Sub:
    iTotal = 0
    For i = 0 To 11
        iTotal = iTotal + Val(txtPercent(i).Text)
    Next
    txtTotalPercent.Text = iTotal
    
    Return

End Sub

Private Sub panelDiff_Click()
    If opt50.Value = True Then opt50.SetFocus
    If opt100.Value = True Then opt100.SetFocus
    If opt500.Value = True Then opt500.SetFocus

End Sub

Private Sub spinPer_SpinDown(Index As Integer)
    Dim iTotal      As String
    
    txtPercent(Index).Text = Val(txtPercent(Index)) - 1
    
    iTotal = 0
    For i = 0 To 11
        iTotal = iTotal + Val(txtPercent(i).Text)
    Next
    txtTotalPercent.Text = iTotal
    

    
End Sub

Private Sub spinPer_SpinUp(Index As Integer)
    Dim iTotal      As String
    
    txtPercent(Index).Text = Val(txtPercent(Index)) + 1
    
    iTotal = 0
    For i = 0 To 11
        iTotal = iTotal + Val(txtPercent(i).Text)
    Next
    txtTotalPercent.Text = iTotal

End Sub

Private Sub txtTitle_Click(Index As Integer)
    
    If opt50.Value = True Then opt50.SetFocus
    If opt100.Value = True Then opt100.SetFocus
    If opt500.Value = True Then opt500.SetFocus

End Sub
