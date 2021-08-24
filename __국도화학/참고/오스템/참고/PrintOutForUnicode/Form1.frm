VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMain 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Print Out For Unicode"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13830
   BeginProperty Font 
      Name            =   "굴림"
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
   ScaleHeight     =   5700
   ScaleWidth      =   13830
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   795
      Left            =   7950
      TabIndex        =   5
      Top             =   4560
      Width           =   1755
   End
   Begin VB.PictureBox Picture1 
      Height          =   2265
      Left            =   10770
      ScaleHeight     =   2205
      ScaleWidth      =   2655
      TabIndex        =   3
      Top             =   2040
      Width           =   2715
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4200
      TabIndex        =   1
      Top             =   1980
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   6000
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   705
      Left            =   2610
      TabIndex        =   4
      Top             =   4710
      Width           =   4125
   End
   Begin MSForms.TextBox TextBox2 
      Height          =   1305
      Left            =   240
      TabIndex        =   2
      Top             =   2610
      Width           =   9105
      VariousPropertyBits=   746604571
      Size            =   "16060;2302"
      FontName        =   "굴림"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   1800
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   10035
      VariousPropertyBits=   -1400616933
      ScrollBars      =   1
      Size            =   "17701;3175"
      FontName        =   "Calibri"
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function TextOutW Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As Long, ByVal nCount As Long) As Long


Private Sub cmdPrint_Click()
    On Error GoTo cancel
    Dim dY As Long
'    comDlg.ShowPrinter
'
'    Printer.FontName = TextBox1.FontName
'    Printer.FontSize = TextBox1.FontSize
    Picture1.FontName = TextBox1.FontName

'    Printer.Print " "
    For dY = 1 To 1
'        Printer.FontSize = 10 + dY * 2
'        Call TextOutW(Printer.hdc, 10, dY * 50, StrPtr(TextBox1), Len(TextBox1))
        Call TextOutW(Picture1.hdc, 1, 1, StrPtr(TextBox1), Len(TextBox1))
'        Call TextOutW(TextBox2, 10, dY * 50, StrPtr(TextBox1), Len(TextBox1))
'        Call TextOutW(Label1, 10, dY * 50, StrPtr(TextBox1), Len(TextBox1))
'
    Next
        
    
'    Printer.EndDoc
cancel:
End Sub
'
'Private Sub Command1_Click()
'
'End Sub
'
'
'Private Sub tbbopenfile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbbopenfile.Click
'        OpenFileText.InitialDirectory = "C:\"
'        OpenFileText.Title = "파일열기"
'        OpenFileText.FileName = ""
'
'        OpenFileText.Filter = "Text Files (:*.txt)|*.txt"
'        OpenFileText.FilterIndex = 1
'
'        If OpenFileText.ShowDialog() = Windows.Forms.DialogResult.OK Then
'            Dim objReader As New System.IO.StreamReader(OpenFileText.FileName)
'            TextBox1.Text = objReader.ReadToEnd
'            objReader.Close()
'
'        End If
'End Sub

Private Sub TextBox1_Change()
'    Dim dY As Long
'
'    TextBox1.Text = TextOutW(Printer.hdc, 10, dY * 50, StrPtr(TextBox1.Text), Len(TextBox1))

End Sub
