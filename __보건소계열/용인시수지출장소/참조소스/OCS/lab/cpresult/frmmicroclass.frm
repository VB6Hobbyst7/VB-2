VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMicroClass 
   Caption         =   "미생물검사 보고단계 확인"
   ClientHeight    =   900
   ClientLeft      =   7140
   ClientTop       =   7590
   ClientWidth     =   4635
   BeginProperty Font 
      Name            =   "굴림체"
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
   ScaleHeight     =   900
   ScaleWidth      =   4635
   Begin VB.Frame Frame1 
      Caption         =   "보고차수확인"
      Height          =   645
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   2670
      Begin VB.OptionButton Option2 
         Caption         =   "최종보고"
         Height          =   285
         Left            =   1395
         TabIndex        =   3
         Top             =   225
         Width           =   1050
      End
      Begin VB.OptionButton Option1 
         Caption         =   "예비보고"
         Height          =   285
         Left            =   180
         TabIndex        =   2
         Top             =   225
         Width           =   1050
      End
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   555
      Left            =   2835
      TabIndex        =   1
      Top             =   180
      Width           =   1635
      Caption         =   "확인"
      PicturePosition =   327683
      Size            =   "2884;979"
      Picture         =   "frmMicroClass.frx":0000
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmMicroClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sJeobsuDt       As String
Dim sSLipno1        As String
Dim sSLipno2        As String
Dim sPtno           As String


Private Sub cmdOk_Click()
    Dim sJeobsuDt       As String
    Dim sSLipno1        As String
    Dim sSLipno2        As String
    Dim sPtno           As String
    
    
    sJeobsuDt = Format(frmResult.dtJeobsu.Value, "yyyy-MM-dd")
    sSLipno1 = Val(Left(frmResult.cmbSLip.Text, 2))
    sSLipno2 = Val(frmResult.txtSLipno2.Text)
    sPtno = frmResult.txtPtno.Text
    
    Dim sStatus     As String
    
    
    If Option1.Value = True Then
        sStatus = "P"   '부분결과(예비보고)
    ElseIf Option2.Value = True Then
        sStatus = "C"   '검사완료(최종보고)
    Else
        sStatus = "R"   '접수중 상태
    End If
    
    GoSub Update_General_Status
    Unload Me
    
    Exit Sub
    
    
Update_General_Status:
    strSql = ""
    strSql = strSql & " UPDATE TWEXAM_General"
    strSql = strSql & " SET    Status    = '" & sStatus & "'"
    strSql = strSql & " WHERE  JeobsuDt  =  TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    SLipno1   =  " & Val(sSLipno1)
    strSql = strSql & " AND    SLipno2   =  " & Val(sSLipno2)
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
        
    Return
    
    
End Sub

Private Sub Form_Load()
    Dim sJeobsuDt       As String
    Dim sSLipno1        As String
    Dim sSLipno2        As String
    Dim sPtno           As String
    
    
    sJeobsuDt = Format(frmResult.dtJeobsu.Value, "yyyy-MM-dd")
    sSLipno1 = Val(Left(frmResult.cmbSLip.Text, 2))
    sSLipno2 = Val(frmResult.txtSLipno2.Text)
    sPtno = frmResult.txtPtno.Text
    
    GoSub GET_General_Status
    Exit Sub
    
GET_General_Status:
    strSql = ""
    strSql = strSql & " SELECT Status"
    strSql = strSql & " FROM   TWEXAM_General"
    strSql = strSql & " WHERE  JeobsuDt  =  TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    SLipno1   =  " & Val(sSLipno1)
    strSql = strSql & " AND    SLipno2   =  " & Val(sSLipno2)
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Select Case adoSet.Fields("Status").Value & ""
        Case "R": Option1.Value = True
        Case "C": Option2.Value = True
        Case "P": Option1.Value = True
        Case "U": Option1.Value = True
    End Select
    Call adoSetClose(adoSet)
    
    Return
    
    
End Sub
