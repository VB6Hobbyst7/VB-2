VERSION 5.00
Begin VB.Form frmServer 
   Caption         =   " ���� ���� "
   ClientHeight    =   3480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9375
   BeginProperty Font 
      Name            =   "���� ���"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6810
      TabIndex        =   6
      Top             =   2610
      Width           =   1125
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7980
      TabIndex        =   5
      Top             =   2610
      Width           =   1125
   End
   Begin VB.TextBox txtUSEURL 
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   1515
      TabIndex        =   4
      Text            =   "127.0.0.1"
      Top             =   540
      Width           =   7605
   End
   Begin VB.TextBox txtAPIURL 
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1500
      TabIndex        =   3
      Text            =   "127.0.0.1"
      Top             =   1020
      Width           =   7605
   End
   Begin VB.TextBox txtSTDURL 
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1500
      TabIndex        =   2
      Text            =   "127.0.0.1"
      Top             =   1530
      Width           =   7605
   End
   Begin VB.OptionButton optURL 
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   210
      TabIndex        =   1
      Top             =   1110
      Value           =   -1  'True
      Width           =   1245
   End
   Begin VB.OptionButton optURL 
      Caption         =   "���߱�"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   210
      TabIndex        =   0
      Top             =   1620
      Width           =   1245
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "��뼭��"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   7
      Top             =   615
      Width           =   915
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdConfirm_Click()
    Dim strUseUrl  As String
    
    On Error GoTo ErrorHandler
    
    If MsgBox("������ �����Ͻðڽ��ϱ�?", vbCritical + vbOKCancel + vbDefaultButton2, "Ȯ��!") = vbCancel Then
        Unload Me
        Exit Sub
    Else
        If optURL(0).Value = True Then
            strUseUrl = "API"   '���
        Else
            strUseUrl = "STD"
        End If
        
        Call WritePrivateProfileString("HOSP", "USEURL", strUseUrl, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("HOSP", "APIURL", txtAPIURL.Text, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("HOSP", "STDURL", txtSTDURL.Text, App.PATH & "\INI\" & gMACH & ".ini")
        
        Unload Me
    End If
        
    Exit Sub
 
ErrorHandler:
    Resume Next
    If MsgBox("��ΰ� ���� �ʽ��ϴ�", vbCritical + vbOKCancel + vbDefaultButton2, "�����ư") = vbCancel Then
        Unload Me
    End If

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    txtAPIURL.Text = gHOSP.APIURL
    txtSTDURL.Text = gHOSP.STDURL
     
    If gHOSP.USEURL = "API" Then
        optURL(0).Value = True
        txtUSEURL.Text = txtAPIURL.Text
    Else
        optURL(1).Value = True
        txtUSEURL.Text = txtSTDURL.Text
    End If
    
End Sub

Private Sub optURL_Click(Index As Integer)
    
    Select Case Index
        Case 0:     txtUSEURL.Text = txtAPIURL.Text
        Case 1:     txtUSEURL.Text = txtSTDURL.Text
    End Select
    
End Sub

