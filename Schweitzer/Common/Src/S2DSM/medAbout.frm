VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '���� ����
   ClientHeight    =   5070
   ClientLeft      =   4110
   ClientTop       =   3090
   ClientWidth     =   5535
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   5535
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  '����
      Height          =   3090
      Left            =   0
      ScaleHeight     =   3090
      ScaleWidth      =   5535
      TabIndex        =   6
      Top             =   0
      Width           =   5535
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "Ver 1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2880
         TabIndex        =   11
         Top             =   1350
         Width           =   720
      End
      Begin VB.Label lblCopyright 
         Alignment       =   2  '��� ����
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "Copyright 1999  Daeryun MTS Co., Ltd."
         Height          =   225
         Index           =   0
         Left            =   1065
         TabIndex        =   10
         Top             =   1950
         Width           =   3150
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "LIS"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   27.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   645
         Index           =   1
         Left            =   1815
         TabIndex        =   9
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label lblHomePage 
         Alignment       =   2  '��� ����
         BackStyle       =   0  '����
         Caption         =   "http://www.medcom.co.kr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1230
         MouseIcon       =   "medAbout.frx":0000
         MousePointer    =   99  '����� ����
         TabIndex        =   8
         Top             =   2460
         Width           =   2715
      End
      Begin VB.Label lblCopyright 
         Alignment       =   2  '��� ����
         BackStyle       =   0  '����
         Caption         =   "Seoul, Korea"
         Height          =   255
         Index           =   1
         Left            =   1245
         TabIndex        =   7
         Top             =   2190
         Width           =   2640
      End
      Begin VB.Image Image2 
         Height          =   960
         Left            =   120
         Picture         =   "medAbout.frx":030A
         Top             =   90
         Width           =   960
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   840
         Picture         =   "medAbout.frx":3554
         Top             =   330
         Width           =   1920
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FCEFE9&
      Caption         =   "Ȯ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4140
      Style           =   1  '�׷���
      TabIndex        =   0
      Top             =   4455
      Width           =   1140
   End
   Begin VB.Label lblAvailMem 
      Alignment       =   2  '��� ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "�ӻ󺴸���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1980
      TabIndex        =   13
      Top             =   4635
      Width           =   1140
   End
   Begin VB.Label lblTotMem 
      Alignment       =   2  '��� ����
      BackColor       =   &H00FFFFFF&
      Caption         =   "�ӻ󺴸���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1725
      TabIndex        =   12
      Top             =   4395
      Width           =   1140
   End
   Begin VB.Label lblMsg 
      BackColor       =   &H00E0E0E0&
      Caption         =   " �� ��ǰ�� ���� ����ڿ��� ����� �㰡�Ǿ����ϴ�."
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   225
      TabIndex        =   5
      Top             =   3330
      Width           =   4995
   End
   Begin VB.Label lblUser 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�ӻ󺴸���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   225
      TabIndex        =   4
      Top             =   3825
      Width           =   1140
   End
   Begin VB.Label lblHospital 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "��õ�ǰ����� �μ� �溴��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   225
      TabIndex        =   3
      Top             =   3615
      Width           =   2100
   End
   Begin VB.Label lblWarning 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "System Resource :"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   4
      Left            =   240
      TabIndex        =   2
      Top             =   4635
      Width           =   1635
   End
   Begin VB.Label lblWarning 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "Total Memory :"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   3
      Left            =   225
      TabIndex        =   1
      Top             =   4395
      Width           =   1305
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   
   Dim tmpTotMem As Long, tmpAvailMem As Long
   
   lblProductName(1).Caption = "LIS"     'App.Comments  ' ������Ʈ��
   lblVersion.Caption = "Version " & App.Major & "." & _
                                    App.Minor & "." & App.Revision    '����
   lblHospital.Caption = HospitalNm
   lblUser.Caption = medGetComNm
   
   Call medSysMem(tmpTotMem, tmpAvailMem)
   lblTotMem.Caption = Format(tmpTotMem / 1024, "###,###,###") & " KB"
   lblAvailMem.Caption = Format((tmpAvailMem / tmpTotMem) * 100, "###") & " %"
   
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHomePage.FontUnderline = False

End Sub

Private Sub lblHomePage_Click()

   Dim i As Double
   Dim MyHomePage As String
   Dim FileName As String
   Dim FileNumber As Integer
   Dim BrowserExec As String * 255
   Dim BrowserExecNm As String

   Dim RetVal As Long


   MyHomePage = lblHomePage.Caption
   
   BrowserExec = Space(255)

   FileName = App.Path & "\temphtm.HTM"
   FileNumber = FreeFile() ' Get unused file number

   '����� �������� Path�� ��Ī�� �˱�����
   '�Ͻ������� HTML������ �����
   '==> ���α׷��� ���κп��� �����Ѵ�
   
   Open FileName For Output As #FileNumber
   Write #FileNumber, " " ' Output text
   Close #FileNumber ' Close file

   ' Then find the application associated with it.

   RetVal = FindExecutable(FileName, Dummy, BrowserExec)
   For i = 1 To Len(BrowserExec)
       If Mid(BrowserExec, i, 1) < " " Then
           Mid(BrowserExec, i, 1) = " "
       End If
   Next i

   BrowserExecNm = Trim$(BrowserExec)
   BrowserExecNm = Mid(BrowserExecNm, 1, InStr(1, BrowserExecNm, " ") - 1)
   
   ' If an application is found, launch it!
   If RetVal <= 32 Or IsEmpty(BrowserExec) Then ' Error
       MsgBox "Could not find a browser"
   Else
       i = Shell(BrowserExecNm & " " & MyHomePage, vbNormalFocus)
   End If
   Kill FileName ' delete temp HTML file

End Sub

Private Sub lblHomePage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHomePage.FontUnderline = True
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHomePage.FontUnderline = False

End Sub
