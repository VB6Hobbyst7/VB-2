VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmRmk 
   Caption         =   "Remark ���� �Է�,����,�����۾�"
   ClientHeight    =   3705
   ClientLeft      =   2145
   ClientTop       =   2955
   ClientWidth     =   7125
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "����ü"
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
   ScaleHeight     =   3705
   ScaleWidth      =   7125
   Begin Threed.SSCommand cmdExit 
      Height          =   915
      Left            =   6180
      TabIndex        =   11
      Top             =   60
      Width           =   795
      _Version        =   65536
      _ExtentX        =   1402
      _ExtentY        =   1614
      _StockProps     =   78
      Caption         =   "����"
      BevelWidth      =   1
      Outline         =   0   'False
      Picture         =   "frmRmk.frx":0000
   End
   Begin Threed.SSCommand cmdDelete 
      Height          =   915
      Left            =   5370
      TabIndex        =   10
      Top             =   60
      Width           =   795
      _Version        =   65536
      _ExtentX        =   1402
      _ExtentY        =   1614
      _StockProps     =   78
      Caption         =   "����Ȯ��"
      BevelWidth      =   1
      Outline         =   0   'False
      Picture         =   "frmRmk.frx":08DA
   End
   Begin Threed.SSCommand cmdInsert 
      Height          =   915
      Left            =   4560
      TabIndex        =   9
      Top             =   60
      Width           =   795
      _Version        =   65536
      _ExtentX        =   1402
      _ExtentY        =   1614
      _StockProps     =   78
      Caption         =   "�Է�Ȯ��"
      BevelWidth      =   1
      Outline         =   0   'False
      Picture         =   "frmRmk.frx":11B4
   End
   Begin Threed.SSPanel panelLen 
      Align           =   2  '�Ʒ� ����
      Height          =   315
      Left            =   0
      TabIndex        =   8
      Top             =   3390
      Width           =   7125
      _Version        =   65536
      _ExtentX        =   12568
      _ExtentY        =   556
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin VB.ListBox lstCode 
      BackColor       =   &H00FFFFC0&
      Height          =   2580
      Left            =   60
      TabIndex        =   7
      Top             =   660
      Width           =   1635
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   435
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   767
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Begin VB.TextBox txtSlipcode 
         Appearance      =   0  '���
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   60
         Width           =   735
      End
      Begin VB.TextBox txtSlipName 
         Appearance      =   0  '���
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   60
         Width           =   3375
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000007&
         Caption         =   "Label1"
         Height          =   255
         Left            =   900
         TabIndex        =   6
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000007&
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.TextBox txtRmkCode 
      Height          =   315
      Left            =   1740
      TabIndex        =   0
      Top             =   660
      Width           =   1635
   End
   Begin VB.TextBox txtRemark 
      Height          =   2235
      Left            =   1740
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1020
      Width           =   5235
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmRmk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Function Select_AbbCode(ByVal sCode As String) As Integer
    
    If Trim(sCode) = "" Then
        Select_AbbCode = False
        Exit Function
    End If
    
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_Remark"
    strSql = strSql & " WHERE  Exgubun  =  '" & Trim(txtSlipCode.Text) & "'"
    strSql = strSql & " Order  by AbbCode"
    If False = adoSetOpen(strSql, adoSet) Then
        Select_AbbCode = False
        Exit Function
    End If
        
    Select_AbbCode = True
    lstCode.Clear
    Do Until adoSet.EOF
        lstCode.AddItem Trim(adoSet.Fields("Abbcode").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)

End Function



Private Sub cmdDelete_Click()
    
    If GetWindowTextLength(txtRemark.hwnd) > 400 Then
        MsgBox "Remark ������ 400 Byte �̻��� �Ǹ� �Է��� �ȵ˴ϴ�"
        txtRemark.SetFocus
        txtRemark.SelStart = 0
        txtRemark.SelLength = Len(txtRemark.Text)
        Exit Sub
    End If
    
    If Trim(txtRmkCode.Text) = "" Then
        MsgBox "������ Remark Code �� �����ϴ�!..", vbOKOnly + vbInformation, "����Miss Information"
        Exit Sub
    End If
    
    If Trim(txtSlipCode.Text) = "" Then
        MsgBox "������ �˻�������  �� �����ϴ�!..", vbOKOnly + vbInformation, "����Miss Information"
        Exit Sub
    End If
    
    If vbNo = MsgBox("�ش� �ڵ带 �����Ͻðڽ��ϱ�?", vbYesNo + vbQuestion, "����Ȯ�� Box") Then
        Exit Sub
    End If
    
    
    strSql = ""
    strSql = strSql & " DELETE"
    strSql = strSql & " FROM   TWEXAM_Remark"
    strSql = strSql & " WHERE  ExGubun  =  '" & Trim(txtSlipCode.Text) & "'"
    strSql = strSql & " AND    AbbCode  =  '" & Trim(txtRmkCode.Text) & "'"
    
    If adoExec(strSql) Then
        MsgBox "�����Ͽ����ϴ�!.", vbOKOnly + vbInformation, "���� Complete Information"
        If False = Select_AbbCode(Trim(txtSlipCode.Text)) Then
            lstCode.Clear
        End If
    Else
        MsgBox "� ������ ���Ͽ� �������� ���Ͽ����ϴ�!", vbOKOnly + vbInformation, "����Miss Information"
        Exit Sub
    End If
    lstCode.SetFocus
        
End Sub

Private Sub cmdExit_Click()
    Unload Me
    
End Sub

Private Sub cmdInsert_Click()
    
    If GetWindowTextLength(txtRemark.hwnd) > 400 Then
        MsgBox "Remark ������ 400 Byte �̻��� �Ǹ� �Է��� �ȵ˴ϴ�"
        txtRemark.SetFocus
        txtRemark.SelStart = 0
        txtRemark.SelLength = Len(txtRemark.Text)
        Exit Sub
    End If
    
    GoSub Text_Data_Null_Check
    
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_Remark"
    strSql = strSql & " WHERE  ExGubun  =  '" & Trim(txtSlipCode.Text) & "'"
    strSql = strSql & " AND    AbbCode  =  '" & Trim(txtRmkCode.Text) & "'"
    If False = adoSetOpen(strSql, adoSet) Then
        GoSub Remark_Data_Put_Sub
    Else
        GoSub Remark_Data_Update_Sub
    End If
    
    If False = Select_AbbCode(Trim(txtSlipCode.Text)) Then
        lstCode.Clear
    End If
    lstCode.SetFocus
    Exit Sub

'/--------------------------------------------------
Text_Data_Null_Check:
    If Trim(txtRmkCode.Text) = "" Then
        MsgBox "�Է��� Remark Code �� �����ϴ�!..", vbOKOnly + vbInformation, "����Miss Information"
        Exit Sub
    End If
    
    If Trim(txtSlipCode.Text) = "" Then
        MsgBox "�Է��� �˻�������  �� �����ϴ�!..", vbOKOnly + vbInformation, "����Miss Information"
        Exit Sub
    End If

    Return
    
Remark_Data_Put_Sub:
    strSql = ""
    strSql = strSql & " INSERT INTO TWEXAM_Remark"
    strSql = strSql & "       (ExGubun, AbbCode, AbbName)"
    strSql = strSql & " VALUES('" & Trim(txtSlipCode.Text) & "',"
    strSql = strSql & "        '" & Trim(txtRmkCode.Text) & "',"
    strSql = strSql & "        '" & Quot_Conv(Trim(txtRemark.Text)) & "')"
    
    If adoExec(strSql) Then
        MsgBox "�Է� �Ͽ����ϴ�!.", vbOKOnly + vbInformation, "�Է� Complete Information"
        'Unload Me
    Else
        MsgBox "� ������ ���Ͽ� �Է����� ���Ͽ����ϴ�!", vbOKOnly + vbInformation, "�Է�Miss Information"
        Exit Sub
    End If
    
    Return
    
Remark_Data_Update_Sub:
    strSql = ""
    strSql = strSql & " UPDATE TWEXAM_Remark"
    strSql = strSql & " SET    Abbname  =  '" & Quot_Conv(Trim(txtRemark.Text)) & "'"
    strSql = strSql & " WHERE  ExGubun  =  '" & Trim(txtSlipCode.Text) & "'"
    strSql = strSql & " AND    Abbcode  =  '" & Trim(txtRmkCode.Text) & "'"
    
    If adoExec(strSql) Then
        MsgBox "���� �Ͽ����ϴ�!.", vbOKOnly + vbInformation, "���� Complete Information"
        'Unload Me
    Else
        MsgBox "� ������ ���Ͽ� �������� ���Ͽ����ϴ�!", vbOKOnly + vbInformation, "����Miss Information"
        Exit Sub
    End If
    
    Return

End Sub

Private Sub Form_Activate()
    
    If Trim(txtRemark.Text) <> "" Then
        txtRemark.SetFocus
    End If
    
End Sub

Private Sub Form_Load()
    
    Me.Width = 7250
    Me.Height = 4400
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    txtSlipCode.Text = frmRemark.txtExgubun.Text
    txtSlipname.Text = frmRemark.txtExName.Text
    
    If Trim(frmRemark.txtAbbCode.Text) <> "" Then
        Me.txtRmkCode.Text = frmRemark.txtAbbCode.Text
        Me.txtRemark.Text = frmRemark.txtAbbName.Text
    End If
    
    If False = Select_AbbCode(Trim(txtSlipCode.Text)) Then
        lstCode.Clear
    End If

End Sub

Private Sub lstCode_DblClick()
    
    txtRmkCode.Text = lstCode.Text
    Call txtRmkCode_LostFocus
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

Private Sub SSCommand1_Click()
    MsgBox Quot_Conv(txtRemark.Text)
End Sub

Private Sub txtRemark_Change()
    Dim nTextLen        As Long
    
    nTextLen = GetWindowTextLength(txtRemark.hwnd)
    panelLen.Caption = Trim(Str(nTextLen)) & " / 400"
    
    If nTextLen > 399 Then
        sMsg = "400 Byte �� ������ �Է��� �ȵ˴ϴ�!, �ٽ� Read �մϴ�"
        If vbOK = MsgBox(sMsg, vbOKOnly + vbCritical, "400 Byte Check") Then
            txtRemark.Text = ""
            Call txtRmkCode_LostFocus
            Exit Sub
        End If
    End If
    
End Sub

Private Sub txtRmkCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    
End Sub

Public Sub txtRmkCode_LostFocus()
    
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_Remark"
    strSql = strSql & " WHERE  ExGubun = '" & Trim(txtSlipCode.Text) & "'"
    strSql = strSql & " AND    AbbCode = '" & Trim(txtRmkCode.Text) & "'"
    If False = adoSetOpen(strSql, adoSet) Then
        txtRemark.Text = ""
    Else
        txtRemark.Text = adoSet.Fields("Abbname").Value & ""
    End If
    
    Call adoSetClose(adoSet)
    
    panelLen.Caption = GetWindowTextLength(txtRemark.hwnd) & " / 400"
    
End Sub
