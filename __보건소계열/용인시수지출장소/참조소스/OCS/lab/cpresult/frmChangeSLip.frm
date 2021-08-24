VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmChangeSLip 
   Caption         =   "SLipSetting"
   ClientHeight    =   2910
   ClientLeft      =   3285
   ClientTop       =   3015
   ClientWidth     =   4890
   BeginProperty Font 
      Name            =   "바탕체"
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
   ScaleHeight     =   2910
   ScaleWidth      =   4890
   Begin Threed.SSFrame SSFrame2 
      Height          =   915
      Left            =   450
      TabIndex        =   5
      Top             =   1080
      Width           =   4065
      _Version        =   65536
      _ExtentX        =   7170
      _ExtentY        =   1614
      _StockProps     =   14
      Caption         =   "Change Set SLip"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "바탕체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cmbSLip 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   405
         Style           =   2  '드롭다운 목록
         TabIndex        =   6
         Top             =   360
         Width           =   3435
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   825
      Left            =   450
      TabIndex        =   2
      Top             =   135
      Width           =   4065
      _Version        =   65536
      _ExtentX        =   7170
      _ExtentY        =   1455
      _StockProps     =   14
      Caption         =   "Current Set SLip"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "바탕체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtSLipC 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   405
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   315
         Width           =   645
      End
      Begin VB.TextBox txtSLipName 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   315
         Width           =   2760
      End
   End
   Begin MSForms.CommandButton cmdExit 
      Height          =   600
      Left            =   3015
      TabIndex        =   1
      Top             =   2115
      Width           =   1500
      Caption         =   "Exit"
      PicturePosition =   327683
      Size            =   "2646;1058"
      Picture         =   "frmChangeSLip.frx":0000
      FontName        =   "바탕체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdChange 
      Height          =   600
      Left            =   1530
      TabIndex        =   0
      Top             =   2115
      Width           =   1500
      Caption         =   "Change OK"
      PicturePosition =   327683
      Size            =   "2646;1058"
      Picture         =   "frmChangeSLip.frx":08DA
      FontName        =   "바탕체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "frmChangeSLip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Text1_Change()

End Sub

Private Sub cmdChange_Click()
    Dim sSLipno1        As String
    
    sSLipno1 = Left(cmbSLip.Text, 2)
    
    Call SaveSetting("CP", "CPRESULT", "SLip", sSLipno1)
    GiExamNumb = sSLipno1
    
    MsgBox "이 P/C 의 기본 검사종목은 " & vbCrLf & _
            cmbSLip.Text & vbCrLf & _
           "로 변경되었습니다!", vbInformation
    
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
    
End Sub

Private Sub Form_Load()

    GoSub SLip_Select
    
    GiExamNumb = Val(GetSetting("CP", "CPRESULT", "SLip"))
    txtSLipC.Text = GiExamNumb
    txtSLipName.Text = GET_SLipname(txtSLipC.Text)
    Call SetComboBox(cmbSLip, GiExamNumb, 2)
    
    
    Exit Sub
    
    

SLip_Select:
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_Specode"
    strSql = strSql & " WHERE  Codegu = '12'"
    strSql = strSql & " AND    Codeky < '90'"
'C    strSql = strSql & " AND    Codeky < '52'"
    strSql = strSql & " ORDER  BY Codeky"
    
    cmbSLip.Clear
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        cmbSLip.AddItem Trim(adoSet.Fields("Codeky").Value & "") & ". " & _
                        Trim(adoSet.Fields("Codenm").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
        
    Return

End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub
