VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBBS803 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "수가 계산 내역 조회"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdQuery 
      BackColor       =   &H00F4F0F2&
      Caption         =   "조회(&Q)"
      Height          =   420
      Left            =   4200
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   8145
      Width           =   1260
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   420
      Left            =   5505
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   8130
      Width           =   1260
   End
   Begin VB.Frame fraInt 
      BackColor       =   &H00DBE6E6&
      Height          =   7905
      Left            =   60
      TabIndex        =   2
      Top             =   180
      Width           =   10665
      Begin VB.TextBox txtPtId 
         Appearance      =   0  '평면
         Height          =   300
         Left            =   5355
         TabIndex        =   3
         Text            =   "7123456"
         Top             =   255
         Width           =   1155
      End
      Begin FPSpread.vaSpread SS 
         Height          =   7155
         Left            =   75
         TabIndex        =   4
         Top             =   690
         Width           =   10515
         _Version        =   196608
         _ExtentX        =   18547
         _ExtentY        =   12621
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShadowColor     =   16515071
         SpreadDesigner  =   "frmBBS803.frx":0000
      End
      Begin MSComCtl2.DTPicker dtpFrDt 
         Height          =   315
         Left            =   1215
         TabIndex        =   5
         Top             =   270
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   115802115
         CurrentDate     =   36838
      End
      Begin MSComCtl2.DTPicker dtpToDt 
         Height          =   315
         Left            =   2895
         TabIndex        =   6
         Top             =   270
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   115802115
         CurrentDate     =   36838
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   315
         Left            =   6525
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   255
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "환자ID"
         Height          =   180
         Index           =   0
         Left            =   4695
         TabIndex        =   10
         Top             =   315
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
         Height          =   180
         Left            =   2700
         TabIndex        =   9
         Top             =   330
         Width           =   135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "처 방 일 자"
         Height          =   180
         Left            =   150
         TabIndex        =   8
         Top             =   330
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmBBS803"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdQuery_Click()
    Dim SSQL    As String
    Dim DrRS    As Recordset
    Dim i       As Long
    Dim j       As Long
    
    
    SSQL = " SELECT * " & _
           " FROM " & T_INT002 & _
           " WHERE orddt between '" & Format(dtpFrDt.Value, "yyyymmdd") & "' AND '" & Format(dtpToDt.Value, "yyyymmdd") & "' "
     
    If txtPtId <> "" Then
        SSQL = SSQL & " AND " & DBW("ptid", txtPtId.Text, 2)
    End If
    
    Call medClearTable(SS)
    
    Set DrRS = New Recordset
    Call DrRS.Open(SSQL, dbconn)
    If DrRS.EOF Then
'        'dbconn.DisplayErrors
        Set DrRS = Nothing
        Exit Sub
    End If
    
    With DrRS
        If .RecordCount > 0 Then
            SS.Row = 0
            For i = 1 To .Fields.Count
                SS.Col = i
                SS.Value = .Fields(i - 1).name
            Next i
            
            For i = 1 To .RecordCount
                SS.Row = i
                For j = 1 To .Fields.Count
                    SS.Col = j
                    If .Fields(j - 1).name = "ptid" Then
                        SS.Value = .Fields(j - 1)
                    Else
                        SS.Value = .Fields(j - 1)
                    End If
                Next j
                .MoveNext
            Next i
        Else
            MsgBox "해당조건의 자료가 없습니다.", vbInformation + vbOKOnly, Me.Caption
            Call medClearTable(SS)
            txtPtId.Text = "": lblPtNm.Caption = ""
        End If
    End With
    Set DrRS = Nothing
End Sub

Private Sub Form_Activate()
    fraInt.Visible = True
    dtpFrDt = DateAdd("d", -1, GetSystemDate)
    dtpToDt = DateAdd("d", 2, GetSystemDate)
End Sub


Private Sub txtPtid_Change()
    If txtPtId.Text = "" Then lblPtNm.Caption = ""
End Sub

Private Sub txtPtid_GotFocus()
    txtPtId.Tag = txtPtId
End Sub

Private Sub txtPtid_KeyPress(KeyAscii As Integer)
    Call medClearTable(SS)
    lblPtNm.Caption = ""
    
    If KeyAscii = vbKeyReturn Then lblPtNm.Caption = getptnm(txtPtId.Text) 'Call Get_Ptnm(txtPtId.Text)
    
End Sub
