VERSION 5.00
Begin VB.Form frmComment 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�ڸ�Ʈ ����"
   ClientHeight    =   8265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9765
   Icon            =   "frmComment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   9765
   StartUpPosition =   1  '������ ���
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   560
      Left            =   2190
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   4590
      Width           =   5000
   End
   Begin VB.TextBox txtPath 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2100
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   6090
      Width           =   5000
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6300
      TabIndex        =   9
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7470
      TabIndex        =   8
      Top             =   7440
      Width           =   1095
   End
   Begin VB.TextBox txtSGPB5 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   560
      Left            =   2430
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   3420
      Width           =   5000
   End
   Begin VB.TextBox txtSGRV16 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   560
      Left            =   2820
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2580
      Width           =   5000
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '�� ����
      BackColor       =   &H00808000&
      BorderStyle     =   0  '����
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   9765
      TabIndex        =   1
      Top             =   0
      Width           =   9765
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   90
         Top             =   90
         Width           =   2865
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackStyle       =   0  '����
         Caption         =   "�ڸ�Ʈ ����"
         BeginProperty Font 
            Name            =   "���� ���"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   12
         Left            =   210
         TabIndex        =   2
         Top             =   180
         Width           =   2625
      End
   End
   Begin VB.TextBox txtHeader 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   6435
   End
   Begin VB.Label Label1 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "INCONCLUSIVE �ڸ�Ʈ"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   3
      Left            =   390
      TabIndex        =   13
      Top             =   3660
      Width           =   1950
   End
   Begin VB.Label Label1 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "FOOTER"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   270
      TabIndex        =   11
      Top             =   6210
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "POSITIVE �ڸ�Ʈ"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   2
      Left            =   360
      TabIndex        =   7
      Top             =   2520
      Width           =   1425
   End
   Begin VB.Label Label1 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "NEGATIVE �ڸ�Ʈ"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   1515
   End
   Begin VB.Label Label1 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "HEADER"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   4
      Left            =   360
      TabIndex        =   3
      Top             =   1290
      Width           =   720
   End
End
Attribute VB_Name = "frmComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private Sub cmdConfirm_Click()
'    Dim strSGCovid19    As String
'    Dim strSDCovid19    As String
'    Dim strSGRV16       As String
'    Dim strSGPB5        As String
'
'
'    If MsgBox("������ �����Ͻðڽ��ϱ�?", vbCritical + vbOKCancel + vbDefaultButton2, "Ȯ��!") = vbCancel Then
'        Unload Me
'        Exit Sub
'    Else
'        strSGCovid19 = Replace(txtSGCovid19.Text, vbCrLf, "CHR(10)CHR(13)")
'        strSGRV16 = Replace(txtSGRV16.Text, vbCrLf, "CHR(10)CHR(13)")
'        strSGPB5 = Replace(txtSGPB5.Text, vbCrLf, "CHR(10)CHR(13)")
'
'        Call WritePrivateProfileString("COMMENT", "SGCOVID", strSGCovid19, App.PATH & "\INI\" & gMACH & ".ini")
'        Call WritePrivateProfileString("COMMENT", "RV16", strSGRV16, App.PATH & "\INI\" & gMACH & ".ini")
'        Call WritePrivateProfileString("COMMENT", "PB5", strSGPB5, App.PATH & "\INI\" & gMACH & ".ini")
'        Call WritePrivateProfileString("COMMENT", "PATH", txtPath.Text, App.PATH & "\INI\" & gMACH & ".ini")
'
'        Unload Me
'    End If
'
'End Sub
'
'Private Sub cmdExit_Click()
'    Unload Me
'End Sub
'
'Private Sub Form_Load()
'    Dim strHeader  As String
'
'    strHeader = ""
'    strHeader = strHeader & "[�˻��] 2019-nCoV, real-time RT PCR" & vbCrLf
'    strHeader = strHeader & "" & vbCrLf
'
'    strFooter = ""
'    strFooter = strFooter & "[�˻���] Real-time RT-PCR" & vbCrLf
'    strFooter = strFooter & "[�˻簳��]" & vbCrLf
'    strFooter = strFooter & ". �����ڿ� �˻� : 2019-nCoV (2019-Novel Coronavirus)�� E segment, RdRP segment�� real-time RT PCR�� ������." & vbCrLf
'    strFooter = strFooter & ". ������ȯ�� ����: �߿�, ��ħ, ȣ����, ����, ����, ������, ȣ�����, ������ ��ũ, �ٹ߼� ��� ���� ���� ������ 2019-nCoV�� ������"
'    strFooter = strFooter & "" & vbCrLf
'    strFooter = strFooter & "--"
'    strFooter = strFooter & "[�˻繮��] ����뺴�� ���ܰ˻����а� ���������˻��, T. (02) 2072-2937, 0883"
'    strFooter = strFooter & "" & vbCrLf
'    strFooter = strFooter & "-------------------------------------------------------------------------" & vbCrLf
'    strFooter = strFooter & "������(�ǵ���) : ������ M.T / �ڼ��� M.D / ���ü� M.D / �ڼ��� M.D / ������ M.D" & vbCrLf
'    strFooter = strFooter & "-------------------------------------------------------------------------" & vbCrLf
'
'        If UCase(pResult) = "POSITIVE" Then
'            strPCmnt = ""
'            strPCmnt = strPCmnt & "[�˻��� �� �ǰ�]" & vbCrLf
'            strPCmnt = strPCmnt & "2019-nCoV:Positive(�缺) (��ü: " & strSpcNm & ")" & vbCrLf
'            strPCmnt = strPCmnt & "(E: " & strEval & ", RdRp: " & strRdRpVal & " )" & vbCrLf
'            strPCmnt = strPCmnt & "" & vbCrLf
'            strPCmnt = strPCmnt & "* Comment: 2020.04.08 1PM ���� �˻� kit ����Ǿ����� �����Ͻñ� �ٶ��ϴ�." & vbCrLf
'            strPCmnt = strPCmnt & "" & vbCrLf
'            strPCmnt = strPCmnt & "--" & vbCrLf
'
'            strCmnt = strPCmnt
'        End If
'
'        If UCase(pResult) = "NEGATIVE" Then
'            strNCmnt = ""
'            strNCmnt = strNCmnt & "[�˻��� �� �ǰ�]" & vbCrLf
'            strNCmnt = strNCmnt & "2019-nCoV:Negative(����) (��ü: " & strSpcNm & ")" & vbCrLf
'            strNCmnt = strNCmnt & "" & vbCrLf
'            strNCmnt = strNCmnt & "--" & vbCrLf
'
'            strCmnt = strNCmnt
'        End If
'
'        If UCase(pResult) = "INCONCLUSIVE" Then
'            strICmnt = ""
'            strICmnt = strICmnt & "[�˻��� �� �ǰ�]" & vbCrLf
'            strICmnt = strICmnt & "2019-nCoV:Inconclusive(�̰���) (��ü: " & strSpcNm & ")" & vbCrLf
'            strICmnt = strICmnt & "(E: " & strEval & ", RdRp: " & strRdRpVal & " )" & vbCrLf
'            strICmnt = strICmnt & "" & vbCrLf
'            If strEval <> "" And strRdRpVal = "" Then
'                strICmnt = strICmnt & "* Comment: E segment�� ���������� Ȯ�εǾ���, RdRp segmentsms�� �����Դϴ�." & vbCrLf
'            ElseIf strEval = "" And strRdRpVal <> "" Then
'                strICmnt = strICmnt & "* Comment: E segment�� �����̳�, RdRp segment�� ���� ������ Ȯ�εǾ����ϴ�." & vbCrLf
'            End If
'            strICmnt = strICmnt & "�߰� Ȯ���� ���Ͽ� ���� ä���� ��ü�� ���ڰ˻�Ƿ� �����ֽø� �˻� �����ϰڽ��ϴ�." & vbCrLf
'            strICmnt = strICmnt & "" & vbCrLf
'            strICmnt = strICmnt & "--" & vbCrLf
'
'            strCmnt = strICmnt
'        End If
'
'
'    txtHeader.Text = strHeader
'
'
'    txtSGCovid19.Text = Replace(gCFXCmnt.SGCOVID, "CHR(10)CHR(13)", vbCrLf)
'    txtSGRV16.Text = Replace(gCFXCmnt.RV16, "CHR(10)CHR(13)", vbCrLf)
'    txtSGPB5.Text = Replace(gCFXCmnt.PB5, "CHR(10)CHR(13)", vbCrLf)
'    txtPath.Text = gCFXCmnt.PATH
'
'End Sub
'
'Private Sub Form_KeyPress(KeyAscii As Integer)
'
'    If KeyAscii = vbKeyEscape Then
'        Unload Me
'    End If
'
'End Sub
'
