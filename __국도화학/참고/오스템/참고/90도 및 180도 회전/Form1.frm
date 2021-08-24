VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5085
   ClientLeft      =   9360
   ClientTop       =   3240
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5085
   ScaleWidth      =   7290
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   390
      Left            =   150
      TabIndex        =   4
      Top             =   1650
      Width           =   1965
   End
   Begin VB.CommandButton Command1 
      Caption         =   "180�� ȸ��"
      Height          =   390
      Index           =   2
      Left            =   150
      TabIndex        =   3
      Top             =   1125
      Width           =   1965
   End
   Begin VB.CommandButton Command1 
      Caption         =   "-90�� ȸ��"
      Height          =   390
      Index           =   1
      Left            =   150
      TabIndex        =   2
      Top             =   675
      Width           =   1965
   End
   Begin VB.CommandButton Command1 
      Caption         =   "90�� ȸ��"
      Height          =   390
      Index           =   0
      Left            =   150
      TabIndex        =   1
      Top             =   225
      Width           =   1965
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '���
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   2325
      ScaleHeight     =   279
      ScaleMode       =   3  '�ȼ�
      ScaleWidth      =   275
      TabIndex        =   0
      Top             =   135
      Width           =   4155
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
    On Error Resume Next
    Dim w1 As Long, h1 As Long 'ũ��
    Dim w2 As Long, h2 As Long 'ũ��
    
    Dim x As Long, y As Long '����
    Dim xx As Long, yy As Long '����
    Dim r As Single, g As Single, b As Single
    
    '���� �̹��� ����
    w1 = Picture1.ScaleWidth
    h1 = Picture1.ScaleHeight
    
    '��Ʈ�� �迭�� �غ��Ѵ�.
    Call SetBitmapArray_Input(w1, h1)
    
    '�̹����� ��Ʈ�� �迭�� �����Ѵ�.
    Call GetDIBits(Picture1.hdc, Picture1.Image, 0, h1, BITMAP_INPUT(0, 0, 0), BITMAP_INFO, DIB_RGB_COLORS)
    
    '��Ʈ�� �迭�� �غ��Ѵ�.
    If Index = 0 Or Index = 1 Then '90��, -90�� ȸ��
        '��� �̹��� ����
        w2 = Picture1.ScaleHeight
        h2 = Picture1.ScaleWidth
        
        Call SetBitmapArray_Output(w2, h2)
    Else '180�� ȸ��
        '��� �̹��� ����
        w2 = Picture1.ScaleWidth
        h2 = Picture1.ScaleHeight
        
        Call SetBitmapArray_Output(w2, h2)
    End If
    
    '�迭 ó��
    For y = 0 To h2 - 1
        For x = 0 To w2 - 1
        
            Select Case Index
            Case 1 '�ð���� 90�� ȸ��
                xx = (h2 - 1) - y
                yy = x
            Case 0 '�ݽð���� 90�� ȸ��
                xx = y
                yy = (w2 - 1) - x
            Case 2 '180�� ȸ��
                xx = (w2 - 1) - x
                yy = (h2 - 1) - y
            End Select
            
            'INPUT, OUTPUT
            BITMAP_OUTPUT(0, x, y) = BITMAP_INPUT(0, xx, yy)
            BITMAP_OUTPUT(1, x, y) = BITMAP_INPUT(1, xx, yy)
            BITMAP_OUTPUT(2, x, y) = BITMAP_INPUT(2, xx, yy)
        Next
    Next
    
    '��� ��ó�ڽ� ũ�⸦ �����Ѵ�.
    Picture1.Move Picture1.Left, Picture1.Top, (w2 + 2) * Screen.TwipsPerPixelX, (h2 + 2) * Screen.TwipsPerPixelY
    
    '��Ʈ�� �迭�� �̹����� �����Ѵ�.
    Call SetDIBits(Picture1.hdc, Picture1.Image, 0, h2, BITMAP_OUTPUT(0, 0, 0), BITMAP_INFO, DIB_RGB_COLORS)
End Sub

'����
Private Sub Command2_Click()
    '�̹��� �ҷ�����
    Picture1.Picture = LoadPicture(App.Path + "\image_08.jpg")
End Sub

Private Sub Form_Load()
    '���� ������� ����� �̵��ϱ�
    Form1.Left = (Screen.Width - Form1.Width) \ 2
    Form1.Top = (Screen.Height - Form1.Height) \ 2

    '�̹��� �ҷ�����
    Picture1.Picture = LoadPicture(App.Path + "\image_08.jpg")
End Sub

