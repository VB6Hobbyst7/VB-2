VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Anato_Dict 
   Caption         =   "������������"
   ClientHeight    =   5010
   ClientLeft      =   75
   ClientTop       =   1905
   ClientWidth     =   11550
   ControlBox      =   0   'False
   Icon            =   "ANATO109.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5010
   ScaleWidth      =   11550
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00C0C0C0&
      Height          =   1476
      Left            =   5340
      ScaleHeight     =   1410
      ScaleWidth      =   2025
      TabIndex        =   17
      Top             =   360
      Width           =   2085
      Begin VB.OptionButton optcode1 
         Caption         =   "�����ڵ�"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   930
         Width           =   1095
      End
      Begin VB.OptionButton optcode1 
         Caption         =   "�����ڵ�"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   525
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00808000&
         BorderStyle     =   1  '���� ����
         Caption         =   "��ȸ����"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   2040
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0C0C0&
      Height          =   2445
      Left            =   5340
      ScaleHeight     =   2385
      ScaleWidth      =   4575
      TabIndex        =   11
      Top             =   2250
      Width           =   4632
      Begin VB.TextBox txtCode 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1224
         MaxLength       =   8
         TabIndex        =   12
         Top             =   660
         Width           =   1100
      End
      Begin VB.TextBox txtCodenm 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   1224
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   1245
         Width           =   3288
      End
      Begin VB.Label Label8 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00808000&
         BorderStyle     =   1  '���� ����
         Caption         =   "�� �� �� �� �� �� �� ��"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   5280
      End
      Begin VB.Label Label2 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "��    ��"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   165
         TabIndex        =   15
         Top             =   705
         Width           =   720
      End
      Begin VB.Label Label4 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "�� �� ��"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   165
         TabIndex        =   13
         Top             =   1335
         Width           =   1050
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0C0C0&
      Height          =   1476
      Left            =   7890
      ScaleHeight     =   1410
      ScaleWidth      =   2025
      TabIndex        =   5
      Top             =   360
      Width           =   2085
      Begin VB.OptionButton optName 
         Caption         =   "���ܸ��"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   945
         Width           =   1230
      End
      Begin VB.OptionButton optCode 
         Caption         =   "�ڵ��"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   495
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00808000&
         BorderStyle     =   1  '���� ����
         Caption         =   "��ȸ����"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   2040
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808000&
      Height          =   4455
      Left            =   10080
      ScaleHeight     =   4395
      ScaleWidth      =   1305
      TabIndex        =   1
      Top             =   270
      Width           =   1365
      Begin Threed.SSCommand cmdExit 
         Height          =   1020
         Left            =   0
         TabIndex        =   3
         Top             =   3360
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   1799
         _StockProps     =   78
         Caption         =   "�� ��"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO109.frx":0442
      End
      Begin Threed.SSCommand cmdView 
         Height          =   1020
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   1799
         _StockProps     =   78
         Caption         =   "�� ȸ"
         ForeColor       =   8388736
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO109.frx":075C
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   1020
         Left            =   0
         TabIndex        =   9
         Top             =   1125
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   1799
         _StockProps     =   78
         Caption         =   "�� ��"
         ForeColor       =   8388736
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO109.frx":0BAE
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   1020
         Left            =   0
         TabIndex        =   10
         Top             =   2235
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   1799
         _StockProps     =   78
         Caption         =   "�� ��"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO109.frx":1000
      End
   End
   Begin VB.ListBox lstDxDict 
      BackColor       =   &H00EBF5EB&
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      Left            =   75
      TabIndex        =   0
      Top             =   720
      Width           =   5145
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00808000&
      BorderStyle     =   1  '���� ����
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   75
      TabIndex        =   4
      Top             =   360
      Width           =   5145
   End
End
Attribute VB_Name = "Anato_Dict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Dim LsCode              As String * 10
    Dim LsName              As String * 60
    Dim LsClass             As String * 2


Private Sub cmdDelete_Click()

    Dim Response            As Integer
    
    If Trim(txtCode.Text) = "" And Trim(txtCodenm.Text) = "" Then Exit Sub
    
    Response = MsgBox("�ڷḦ �����ұ��?", vbYesNo + vbQuestion + vbDefaultButton2, "���ܺ���")
  
    If Response = vbNo Then Exit Sub
    
    strSQL = " DELETE FROM TWANAT_DICT WHERE Code = '" & Trim(txtCode.Text) & "'"
    
    adoConnect.BeginTrans
    
    Result = AdoExecute(strSQL)
    
    If Result = True And Rowindicator > 0 Then
        adoConnect.CommitTrans
        MsgBox "���� �Ϸ�Ǿ����ϴ�.", vbInformation, "���ܺ�����"
        txtCode.Text = ""
        txtCodenm.Text = ""
    Else
        adoConnect.RollbackTrans
        MsgBox "�۾����� ����ġ ���� ������ �߻��߽��ϴ�.", vbCritical, "����"
    End If
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
    
End Sub


Private Sub cmdSave_Click()
    
    If Trim(txtCode.Text) = "" Then Exit Sub
    
    '��ȸ
    strSQL = ""
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & "   FROM TWANAT_Dict "
    strSQL = strSQL & "  WHERE  CODE   = '" & Trim(txtCode.Text) & "'"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then
        'INSERT
    
        strSQL = ""
        strSQL = strSQL & " INSERT INTO TWANAT_DICT "
        strSQL = strSQL & "  ( CODE, DXDICT ) "
        strSQL = strSQL & "VALUES ('" & Trim(txtCode.Text) & "',"
        strSQL = strSQL & "        '" & Trim(txtCodenm.Text) & "')"
        
        adoConnect.BeginTrans
    
        Result = AdoExecute(strSQL)
        
        If Result = True And Rowindicator > 0 Then
            adoConnect.CommitTrans
            MsgBox "���� �Ϸ�Ǿ����ϴ�.", vbInformation, "���ܺ�����"
        Else
            adoConnect.RollbackTrans
            MsgBox "�۾����� ����ġ ���� ������ �߻��߽��ϴ�.", vbCritical, "����"
        End If
    
    Else
        'UPDATE
        strSQL = ""
        strSQL = strSQL & " UPDATE TWANAT_DICT "
        strSQL = strSQL & " SET    DXDICT = '" & Trim(txtCodenm.Text) & "' "
        strSQL = strSQL & " WHERE  CODE   = '" & Trim(txtCode.Text) & "' "
        
        adoConnect.BeginTrans
        
        Result = AdoExecute(strSQL)
        
        If Result = True And Rowindicator > 0 Then
            adoConnect.CommitTrans
            MsgBox "���� �Ϸ�Ǿ����ϴ�.", vbInformation, "���ܺ�����"
        Else
            adoConnect.RollbackTrans
            MsgBox "�۾����� ����ġ ���� ������ �߻��߽��ϴ�.", vbCritical, "����"
        End If
    
    End If
    
    
'    Call Form_Load
'    txtItemCD.SetFocus
        
    Exit Sub



End Sub

Private Sub cmdView_Click()

    Dim i                   As Integer
    Dim LsTitCode           As String * 10
    Dim LsTitName           As String * 54
        
    txtCode.Text = ""
    txtCodenm.Text = ""
    
    
    LsTitCode = "�� �� ��"
    LsTitName = "�� �� �� �� ��"
    
    If optCode = True Then
        lblTitle = LsTitCode & LsTitName
        strSQL = ""
        strSQL = strSQL & " SELECT * "
        strSQL = strSQL & "   FROM TWANAT_Dict "
        If optcode1(0).Value = True Then
            strSQL = strSQL & "  WHERE SUBSTR(CODE,1,1)  = 'M' "
        Else
            strSQL = strSQL & "  WHERE SUBSTR(CODE,1,1)  = 'T' "
        End If
        strSQL = strSQL & "  ORDER BY Code"
    Else
        lblTitle = LsTitName & LsTitCode
        strSQL = ""
        strSQL = strSQL & " SELECT * "
        strSQL = strSQL & "   FROM TWANAT_Dict "
        If optcode1(0).Value = True Then
            strSQL = strSQL & "  WHERE SUBSTR(CODE,1,1)  = 'M' "
        Else
            strSQL = strSQL & "  WHERE SUBSTR(CODE,1,1)  = 'T' "
        End If
        strSQL = strSQL & "  ORDER BY Dxdict"
    End If
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    lstDxDict.Clear
    Do Until rs.EOF
        LsCode = rs.Fields("Code").Value & ""
        LsName = rs.Fields("DxDict").Value & ""
        
        If optCode = True Then
            lstDxDict.AddItem LsCode & LsName
        Else
            lstDxDict.AddItem LsName & LsCode
        End If
        rs.MoveNext
    Loop
    AdoCloseSet rs
    
End Sub

Private Sub Form_Load()
    optcode1(0).Value = True

End Sub

Private Sub lstDxDict_Click()
    
    If optCode = True Then
         LsCode = Mid(lstDxDict.List(lstDxDict.ListIndex), 1, 10)
         LsName = Mid(lstDxDict.List(lstDxDict.ListIndex), 11, 60)
    Else
         LsCode = Mid(lstDxDict.List(lstDxDict.ListIndex), 11, 60)
         LsName = Mid(lstDxDict.List(lstDxDict.ListIndex), 61, 10)
    End If
  
       txtCode.Text = LsCode
       txtCodenm.Text = LsName

End Sub


Private Sub txtCode_GotFocus()
    
    If txtCode.Text <> "" Then
        MsgBox " �ڵ�� ���� �� �� �����ϴ�. " & vbCrLf & vbCrLf & _
               " �ڵ带 �����ϴ� ���� �ű��ڵ带 ����ϴ� �Ͱ� �����ϴ�." & vbCrLf & vbCrLf & _
               " ���� �ڵ尡 �ʿ� ������ �����Ͻʽÿ�."
    End If
    
    txtCode.SelStart = 0
    txtCode.SelLength = Len(txtCode.Text)

End Sub

Private Sub txtCode_LostFocus()
    txtCode.Text = UCase(txtCode.Text)

End Sub


Private Sub txtCodenm_GotFocus()
    txtCodenm.SelStart = 0
    txtCodenm.SelLength = Len(txtCodenm.Text)

End Sub
