VERSION 5.00
Object = "{B9289CFA-A412-11D4-8C41-00E09878E6B5}#21.0#0"; "Barcod.ocx"
Begin VB.Form Form1 
   Appearance      =   0  '���
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '���� ����
   Caption         =   "(��)����Ƽ���̾�Ƽ  [���ڵ��μ�] v1.0.0"
   ClientHeight    =   6975
   ClientLeft      =   1260
   ClientTop       =   2145
   ClientWidth     =   11235
   BeginProperty Font 
      Name            =   "����"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6975
   ScaleWidth      =   11235
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.Frame Frame3 
      Caption         =   "EAN-8"
      Height          =   2205
      Left            =   180
      TabIndex        =   13
      Top             =   4560
      Width           =   7635
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5820
         TabIndex        =   17
         Top             =   720
         Width           =   555
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Check Digit ���"
         Height          =   495
         Left            =   3960
         TabIndex        =   16
         Top             =   720
         Width           =   1755
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   270
         TabIndex        =   15
         Top             =   720
         Width           =   3135
      End
      Begin VB.CommandButton Command5 
         Caption         =   "���ڵ� ����"
         Height          =   495
         Left            =   5820
         TabIndex        =   14
         Top             =   1470
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "�μ��� ���ڵ�(7�ڸ� �Է�)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   270
         TabIndex        =   18
         Top             =   480
         Width           =   2985
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "EAN-13"
      Height          =   2205
      Left            =   180
      TabIndex        =   6
      Top             =   2250
      Width           =   7635
      Begin VB.CommandButton Command4 
         Caption         =   "���ڵ� ����"
         Height          =   495
         Left            =   5820
         TabIndex        =   11
         Top             =   1470
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   270
         TabIndex        =   9
         Top             =   720
         Width           =   3135
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Check Digit ���"
         Height          =   495
         Left            =   3960
         TabIndex        =   8
         Top             =   720
         Width           =   1755
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5820
         TabIndex        =   7
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "�μ��� ���ڵ�(12�ڸ� �Է�)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   270
         TabIndex        =   10
         Top             =   480
         Width           =   2985
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "���ڵ� ����"
      Height          =   2055
      Left            =   180
      TabIndex        =   1
      Top             =   90
      Width           =   7635
      Begin VB.CommandButton Command1 
         Appearance      =   0  '���
         BackColor       =   &H80000005&
         Caption         =   "���ڵ� �μ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5820
         TabIndex        =   12
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "��   ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5820
         TabIndex        =   5
         Top             =   420
         Width           =   1455
      End
      Begin Barcod.Barcode Barcode1 
         Height          =   585
         Left            =   360
         TabIndex        =   2
         Top             =   750
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "88006611"
         CodeStyle       =   "Code 3 of 9"
      End
      Begin VB.Label lbBarcod 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   4
         Top             =   1410
         Width           =   3225
      End
      Begin VB.Label lbName 
         BackColor       =   &H00FF80FF&
         Caption         =   "Barcode Print"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   3225
      End
      Begin VB.Image Image1 
         Height          =   1575
         Left            =   240
         Picture         =   "BARCODS1.frx":0000
         Top             =   270
         Width           =   3450
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '���
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   3195
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3225
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Command5_Click()

    Barcode1.CodeStyle = "EAN-8"
       
    Barcode1.Caption = Text3.Text & Text4.Text
    
    lbBarcod.Caption = Text3.Text & Text4.Text
    
End Sub

Private Sub Command6_Click()

    Text4.Text = Check_Digit(Text3.Text)
    
End Sub

Private Sub Form_Load()


'   Combo1.AddItem "EAN-8"
'   Combo1.AddItem "EAN-13"
'   Combo1.AddItem "UPC-A"
'   Combo1.AddItem "Code 3 of 9"


End Sub


Private Sub Command1_Click()
    
           
    Picture1.Height = Barcode1.Height
    Picture1.Width = Barcode1.Width
    Picture1.Picture = Barcode1.Image
    Picture1.Refresh
    
    Clipboard.Clear
    Clipboard.SetData Picture1.Image  '���Ĺڽ��� �̹����� Ŭ������� �̵�
        
    Form2.Picture1 = Clipboard.GetData(2)  'Ŭ������ �̹����� ��2�� ���Ĺڽ��� �̵�
        
    Form2.PrintForm
    Unload Form2
    
    
End Sub

Private Sub Command2_Click()

    End
    
End Sub

'=================================================================
'EAN-8 �ڵ忡�� ������ üũ ����Ʈ�� ��� �ϴ� �Լ�
'�����ϸ� üũ����Ʈ ���� ��ȯ�ϰ�
'�����ϸ� "error" ��Ʈ���� ��ȯ�Ѵ�
'�μ� ��Ʈ�� �ڸ����� 7�ڸ��� ���ѵȴ�
'=================================================================
Private Function Check_Digit(ByVal SevenCode As String) As String
    Dim JakHap, HolHap, SumHap, n, K As Integer
    On Error GoTo ErrorHandler
    
    JakHap = 0
    HolHap = 0
    SumHap = 0

    If Len(SevenCode) = 7 Then

        For n = 1 To 7 Step 2
            JakHap = JakHap + Val(Mid(SevenCode, n, 1))
        Next n
        JakHap = JakHap * 3
        For K = 2 To 6 Step 2
            HolHap = HolHap + Val(Mid(SevenCode, K, 1))
        Next K
        SumHap = HolHap + JakHap
        If Val(Right(SumHap, 1)) = 0 Then
            Check_Digit = "0"
            Exit Function
        Else
            Check_Digit = 10 - Val(Right(SumHap, 1))
            Exit Function
        End If
        
    ElseIf Len(SevenCode) = 12 Then
    
        For n = 1 To 11 Step 2
            JakHap = JakHap + Val(Mid(SevenCode, n, 1))
        Next n
        JakHap = JakHap * 3
        For K = 2 To 12 Step 2
            HolHap = HolHap + Val(Mid(SevenCode, K, 1))
        Next K
        SumHap = HolHap + JakHap
        If Val(Right(SumHap, 1)) = 0 Then
            Check_Digit = "0"
            Exit Function
        Else
            Check_Digit = 10 - Val(Right(SumHap, 1))
            Exit Function
        End If
        
    Else
    
        Check_Digit = Error
        
    End If
    
    
    Exit Function
ErrorHandler:
    Check_Digit = "error"
End Function

Private Sub Command3_Click()
   Text2.Text = Check_Digit(Text1.Text)
End Sub

Private Sub Command4_Click()

    Barcode1.CodeStyle = "Code128B"
       
    Barcode1.Caption = Text1.Text & Text2.Text
    
    lbBarcod.Caption = Text1.Text & Text2.Text
    
End Sub






