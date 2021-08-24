VERSION 5.00
Begin VB.Form frmLogView 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�α׺���"
   ClientHeight    =   9825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9240
   Icon            =   "frmLogView.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ݱ�"
      Height          =   405
      Left            =   7530
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtDest 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtSrc 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "���ں�ȯ"
      Height          =   405
      Left            =   4590
      TabIndex        =   2
      Top             =   120
      Width           =   1035
   End
   Begin VB.CommandButton cmdReload 
      Caption         =   "�ٽ� �ҷ�����"
      Height          =   405
      Left            =   150
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtLog 
      BackColor       =   &H00EBFBFF&
      Height          =   9045
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   8955
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '����
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3240
      TabIndex        =   6
      Top             =   210
      Width           =   375
   End
End
Attribute VB_Name = "frmLogView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdReload_Click()

    Call LogLoad

End Sub

Private Sub cmdReplace_Click()

    txtLog.Text = Replace(txtLog.Text, txtSrc.Text, txtDest.Text)
    
End Sub

Private Sub Form_Load()

    Call LogLoad

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub LogLoad()
    Dim FilNum
    Dim sFileName   As String
    Dim FindFile    As String
    Dim TextLine    As String
    Dim strBuffer   As String
    
On Error GoTo RST
    
    FilNum = FreeFile
    
    sFileName = gKUKDO.MACHNM & "_" & Format(CDate(Now), "yyyy-mm-dd")
    
    FindFile = App.PATH & "\Log\" & sFileName & ".txt"
    
    If FindFile <> "" Then
    
        Open App.PATH & "\Log\" & sFileName & ".txt" For Input As FilNum
        
        Do While Not EOF(1)         ' ������ ���� ���� ������ �ݺ��մϴ�.
            Line Input #1, TextLine ' ������ ������ ���� �о���Դϴ�.
            strBuffer = strBuffer & TextLine
        Loop
    
    End If
    
    txtLog.Text = strBuffer
    Close FilNum

Exit Sub
RST:
    
End Sub
