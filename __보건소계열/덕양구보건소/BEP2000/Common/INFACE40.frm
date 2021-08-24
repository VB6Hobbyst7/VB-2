VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form INTface40 
   BorderStyle     =   0  '없음
   Caption         =   "해당일의 검사결과 받기"
   ClientHeight    =   1245
   ClientLeft      =   2715
   ClientTop       =   1920
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1245
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  '아래 맞춤
      Height          =   1260
      Left            =   0
      TabIndex        =   1
      Top             =   -15
      Width           =   5865
      _Version        =   65536
      _ExtentX        =   10345
      _ExtentY        =   2222
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      BorderWidth     =   1
      BevelInner      =   1
      Alignment       =   6
      Begin VB.TextBox txtmmdd 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2655
         MaxLength       =   4
         TabIndex        =   0
         Top             =   435
         Width           =   615
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   510
         Left            =   570
         TabIndex        =   2
         Top             =   360
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   900
         _StockProps     =   15
         Caption         =   "월일입력"
         ForeColor       =   65535
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdcancel 
         Height          =   870
         Index           =   1
         Left            =   4800
         TabIndex        =   4
         Top             =   195
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   1535
         _StockProps     =   78
         Caption         =   "취   소"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Picture         =   "INFACE40.frx":0000
      End
      Begin Threed.SSCommand cmdensure 
         Height          =   870
         Index           =   0
         Left            =   3970
         TabIndex        =   5
         Top             =   195
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   1535
         _StockProps     =   78
         Caption         =   "확   인"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Picture         =   "INFACE40.frx":0E5E
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "####"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1905
         TabIndex        =   3
         Top             =   480
         Width           =   675
      End
   End
End
Attribute VB_Name = "INTface40"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Errkey As Integer
Dim DBopenkey As Integer

Private Sub cmdcancel_Click(Index As Integer)
    
    Unload Me
    FrmFlag = 0
End Sub


Private Sub cmdensure_Click(Index As Integer)

Dim RepairKey As Integer

On Error GoTo repairdb1
        If IsDate(Right$(Format(Now, "yyyy"), 2) & "-" & Left$(txtmmdd, 2) & "-" & Right$(txtmmdd, 2)) = False Then
            MsgBox "날짜입력을 정확히 해 주세요!!"
            txtmmdd.SetFocus
        Else
            strmmdd = machinit & txtmmdd.Text
            textmmdd = txtmmdd.Text
            
            Errkey = False
            Screen.MousePointer = 11
            Call CreateOrOpen_db(strmmdd)
            
            If Errkey = False Then
                DBopenkey = True
            End If
            
            Unload Me
        End If
        
        Screen.MousePointer = 0

Exit Sub

repairdb1:
    
    Errkey = True
    
    If Err = 3049 Then
        MsgBox "데이타가 손상되어 있습니다. 확인을 누르시면 데이타를 복구합니다."
        RepairDatabase (filename & "comm\" & strmmdd)
        Set Db = OpenDatabase(filename & "comm\" & strmmdd)
        RepairKey = True
    Else
        MsgBox Error(Err), vbCritical, Me.Caption
    End If
   
    Resume Next
    
End Sub


Private Sub Form_Load()
            
    'form을 가운데에 위치
    Me.Top = (INTmain00.Height - INTmain00.pnlMain.Height - Me.Height) / 3
    Me.Left = (INTmain00.Width - Me.Width) / 2
    txtmmdd.Text = Format(month(Now), "00") & Format(day(Now), "00")
    Label2.Caption = ":" & machinit
    
    DBopenkey = False
    
    FrmFlag = 40
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If (ImgClickkey = True And Errkey = True) Or (ImgClickkey = True And DBopenkey = False) Then
        ImgClickkey = False
        Exit Sub
    End If
        
    Dim failcomm As Integer
    
    failcomm = False
    
    On Error GoTo frmloaderr
    
    If Len(txtmmdd) = 4 Then
        Load interfacfrm
        
        If failcomm = False Then
            INTface40.Show
        Else
            GoTo frmloaderr
        End If
    End If
               
frmloaderr:
        failcomm = True
        Resume Next

End Sub

Private Sub txtmmdd_GotFocus()

    Call txbox_highlight(txtmmdd)
    
End Sub

Private Sub txtmmdd_KeyDown(KeyCode As Integer, Shift As Integer)

Dim RepairKey As Integer

On Error GoTo repairdb1
    If KeyCode = 13 Then
        If IsDate(Right$(Format(Now, "yyyy"), 2) & "-" & Left$(txtmmdd, 2) & "-" & Right$(txtmmdd, 2)) = False Then
            MsgBox "날짜입력을 정확히 해 주세요!!"
            Call txbox_highlight(txtmmdd)
        Else
            Me.MousePointer = 11
            strmmdd = machinit & txtmmdd.Text
            textmmdd = txtmmdd.Text
            Errkey = False
            
            Screen.MousePointer = 11
            CreateOrOpen_db strmmdd
            
            If Errkey = False Then
                DBopenkey = True
            End If
            
            Unload INTface40
        End If
        Screen.MousePointer = 0
    End If
    
Exit Sub

repairdb1:
    
    Errkey = True
    
    If Err = 3049 Then
        MsgBox "데이타가 손상되어 있습니다. 확인을 누르시면 데이타를 복구합니다."
        RepairDatabase (filename & "comm\" & strmmdd)
        Set Db = OpenDatabase(filename & "comm\" & strmmdd)
        RepairKey = True
    End If
    
    Resume Next

End Sub


