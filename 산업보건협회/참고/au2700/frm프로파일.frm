VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frm�������� 
   BorderStyle     =   1  '���� ����
   Caption         =   "��������"
   ClientHeight    =   8535
   ClientLeft      =   4245
   ClientTop       =   1170
   ClientWidth     =   11955
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
   ScaleHeight     =   8535
   ScaleWidth      =   11955
   Begin VB.Frame Frame1 
      Caption         =   "[����������Ȳ]"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Index           =   1
      Left            =   60
      TabIndex        =   12
      Top             =   780
      Width           =   5535
      Begin FPSpread.vaSpread vas�������� 
         Height          =   7395
         Left            =   60
         TabIndex        =   0
         Top             =   240
         Width           =   5385
         _Version        =   393216
         _ExtentX        =   9499
         _ExtentY        =   13044
         _StockProps     =   64
         ColHeaderDisplay=   0
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   2
         OperationMode   =   2
         SelectBlockOptions=   0
         SpreadDesigner  =   "frm��������.frx":0000
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "[�������ϳ���]"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Index           =   0
      Left            =   5640
      TabIndex        =   9
      Top             =   780
      Width           =   6255
      Begin VB.CommandButton cmdCancel 
         Caption         =   "���(&C)"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   5460
         TabIndex        =   6
         Top             =   300
         Width           =   735
      End
      Begin VB.TextBox txtEQ_PROFILENM 
         Height          =   315
         Left            =   1500
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   660
         Width           =   2355
      End
      Begin VB.TextBox txtEQ_PROFILECD 
         Height          =   315
         Left            =   1500
         MaxLength       =   5
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   300
         Width           =   2355
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "����(&D)"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   4680
         TabIndex        =   5
         Top             =   300
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "����(&S)"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   3900
         TabIndex        =   4
         Top             =   300
         Width           =   735
      End
      Begin FPSpread.vaSpread vas�˻��ڵ� 
         Height          =   6555
         Left            =   60
         TabIndex        =   3
         Top             =   1080
         Width           =   6105
         _Version        =   393216
         _ExtentX        =   10769
         _ExtentY        =   11562
         _StockProps     =   64
         ColHeaderDisplay=   0
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   4
         SpreadDesigner  =   "frm��������.frx":1927
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "���������ڵ�"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�������ϸ�"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   10
         Top             =   720
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "�ݱ�(&Q)"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10860
      TabIndex        =   7
      Top             =   60
      Width           =   1035
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   2
      X1              =   60
      X2              =   11880
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '�������� ����
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '�ܻ�
      Height          =   435
      Index           =   1
      Left            =   3960
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '�������� ����
      FillColor       =   &H000000FF&
      FillStyle       =   0  '�ܻ�
      Height          =   435
      Index           =   0
      Left            =   60
      Top             =   60
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�������� ����"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   9
      Left            =   480
      TabIndex        =   8
      Top             =   180
      Width           =   1890
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FF8080&
      FillStyle       =   0  '�ܻ�
      Height          =   615
      Index           =   1
      Left            =   60
      Shape           =   4  '�ձ� �簢��
      Top             =   60
      Width           =   4035
   End
End
Attribute VB_Name = "frm��������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngMeHeight     As Long '/Me.Height�� �ʱⰪ
Dim lngMeWidth      As Long '/Me.Width�� �ʱⰪ

Private Type ConWHERE   ' ����� ���� ������ ����ϴ�.
   Nm       As String
   Left     As Long
   Top      As Long
   Width    As Long
   Height   As Long
End Type
Dim CW()    As ConWHERE

Public Sub MM_CANCEL()
    If vas��������.MaxRows > 0 Then vas��������.MaxRows = 0
    
    Call MM_KEY_CLEAR
End Sub

Public Sub MM_INITIAL()
    '/Resize�� ���� �ʱ� Size Setting----------------------------------------------------------------------------------------------------/
    For intX = 0 To Me.Count - 1
        Select Case True
            Case TypeOf Me.Controls(intX) Is Menu
            Case TypeOf Me.Controls(intX) Is Line
            Case TypeOf Me.Controls(intX) Is MSComm
            Case Else
                ReDim Preserve CW(intX)
                
                CW(intX).Nm = Me.Controls(intX).Name
                CW(intX).Left = Me.Controls(intX).Left
                CW(intX).Top = Me.Controls(intX).Top
                CW(intX).Width = Me.Controls(intX).Width
                CW(intX).Height = Me.Controls(intX).Height
        End Select
    Next intX
    
    '/Form Size Setting
    lngMeHeight = 9015
    lngMeWidth = 12075
    
    Me.Height = lngMeHeight
    Me.Width = lngMeWidth
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    '/Resize�� ���� �ʱ� Size Setting----------------------------------------------------------------------------------------------------/
    
    '/�ʱ� �ڷ� Setting----------------------------------------------------------------------------------------------------/
    GoSub ADD_ITEM
    '/�ʱ� �ڷ� Setting----------------------------------------------------------------------------------------------------/
    
    '/���� ��Ʈ�� �ʱ�ȭ----------------------------------------------------------------------------------------------------/
    Call MM_CANCEL
    '/���� ��Ʈ�� �ʱ�ȭ----------------------------------------------------------------------------------------------------/
    
'    Call SET_CBO_DT_L(GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_EQUIP, REG_PG_WORKLIST), cboPG_WORKLIST, 1)
'    txtWaitTime = "10"
'    Call SET_CBO_DT_L(GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_EQUIP, REG_PG_QC), cboPG_QC, 1)
    
Exit Sub

'/========================================================================================================================================================================================================/

ADD_ITEM:
    Dim lngRow   As Long
    
    If OpenDB(gtypREG_INFO.DB_CONSTR_LOCAL) = False Then Exit Sub

    gstrQuy = "SELECT * "
    gstrQuy = gstrQuy & vbCrLf & "  FROM EQUIPEXAM "
    gstrQuy = gstrQuy & vbCrLf & " ORDER BY SEQNO "
    If ReadSQL(gstrQuy, ADR) = False Then
        Call CloseDB
        Exit Sub
    End If
    
    If Not ADR Is Nothing Then
        vas�˻��ڵ�.MaxRows = ARC
        
        Do Until ADR.EOF
            lngRow = lngRow + 1
        
            Call SET_CELL(vas�˻��ڵ�, 2, lngRow, Trim(ADR!EQUIPCODE & ""))
            Call SET_CELL(vas�˻��ڵ�, 3, lngRow, Trim(ADR!ExamCode & ""))
            Call SET_CELL(vas�˻��ڵ�, 4, lngRow, Trim(ADR!examname & ""))
            
            ADR.MoveNext
        Loop
        ADR.Close: Set ADR = Nothing
    End If
    Call CloseDB
Return
End Sub

Private Sub MM_KEY_CLEAR()
    txtEQ_PROFILECD = ""
    txtEQ_PROFILENM = ""
    
    vas�˻��ڵ�.ClearRange 1, -1, 1, -1, -1
    
    cmdDelete.Enabled = False
End Sub

Public Function MM_SAVE() As Boolean

End Function

Public Sub MM_VIEW()
    Dim lngRow  As Long
    
    If vas��������.MaxRows > 0 Then vas��������.MaxRows = 0
    
    If OpenDB(gtypREG_INFO.DB_CONSTR_LOCAL) = False Then Exit Sub

    gstrQuy = "SELECT EQ_PROFILECD, EQ_PROFILENM "
    gstrQuy = gstrQuy & vbCrLf & "  FROM EQ_PROFILE "
    gstrQuy = gstrQuy & vbCrLf & " GROUP BY EQ_PROFILECD, EQ_PROFILENM "
    gstrQuy = gstrQuy & vbCrLf & " ORDER BY EQ_PROFILECD "
    If ReadSQL(gstrQuy, ADR) = False Then
        Call CloseDB
        Exit Sub
    End If
    
    If Not ADR Is Nothing Then
        vas��������.MaxRows = ARC
        
        Do Until ADR.EOF
            lngRow = lngRow + 1
        
            Call SET_CELL(vas��������, 1, lngRow, Trim(ADR!EQ_PROFILECD & ""))
            Call SET_CELL(vas��������, 2, lngRow, Trim(ADR!EQ_PROFILENM & ""))
            
            ADR.MoveNext
        Loop
        ADR.Close: Set ADR = Nothing
    End If
    Call CloseDB
End Sub

Private Sub cmdCancel_Click()
    Call MM_KEY_CLEAR
End Sub

Private Sub cmdDelete_Click()
    If Trim(txtEQ_PROFILECD) = "" Then MsgBox "���������ڵ带 (��)�����Ͻʽÿ�!", vbCritical, "���� �Ұ�": txtEQ_PROFILECD.SetFocus: Exit Sub
    
    If MsgBox("�����ϰڽ��ϱ�?", vbQuestion + vbOKCancel, "���� ����") = vbCancel Then Exit Sub
              
    If OpenDB(gtypREG_INFO.DB_CONSTR_LOCAL) = False Then Exit Sub
    
    ADC.BeginTrans

    gstrQuy = "DELETE FROM EQ_PROFILE "
    gstrQuy = gstrQuy & vbCrLf & " WHERE EQ_PROFILECD   = '" & Trim(txtEQ_PROFILECD) & "' "
    If RunSQL(gstrQuy) = False Then
        ADC.RollbackTrans
        Call CloseDB
        Call ErrQuery(gstrQuy, 0)
        Exit Sub
    End If
    
    ADC.CommitTrans
    
    Call CloseDB
    
    Call MM_KEY_CLEAR
    Call MM_VIEW '/���¸� �ٽ� ���̱� ����...
    
    MsgBox "���� �Ǿ����ϴ�.", vbInformation, "Ȯ��"
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim strSaveYN   As String
    
    '/�Է��׸�Check
    If Trim(txtEQ_PROFILECD) = "" Then MsgBox "���������ڵ带 (��)�����Ͻʽÿ�!", vbCritical, "���� �Ұ�": txtEQ_PROFILECD.SetFocus: Exit Sub
    If Trim(txtEQ_PROFILENM) = "" Then MsgBox "�������ϸ��� (��)�����Ͻʽÿ�!", vbCritical, "���� �Ұ�": txtEQ_PROFILENM.SetFocus: Exit Sub
    For intX = 1 To vas�˻��ڵ�.MaxRows
        If GET_CELL(vas�˻��ڵ�, 1, intX) = "1" Then strSaveYN = "Y": Exit For
    Next intX
    If strSaveYN <> "Y" Then MsgBox "������ �˻��ڵ带 �����Ͻʽÿ�!", vbCritical, "���� �Ұ�": Exit Sub
    
    If MsgBox("�����ϰڽ��ϱ�?", vbQuestion + vbOKCancel, "���� ����") = vbCancel Then Exit Sub
              
    If OpenDB(gtypREG_INFO.DB_CONSTR_LOCAL) = False Then Exit Sub
    
    ADC.BeginTrans
    
    gstrQuy = "DELETE FROM EQ_PROFILE "
    gstrQuy = gstrQuy & vbCrLf & " WHERE EQ_PROFILECD   = '" & Trim(txtEQ_PROFILECD) & "' "
    If RunSQL(gstrQuy) = False Then
        ADC.RollbackTrans
        Call CloseDB
        Call ErrQuery(gstrQuy, 0)
        Exit Sub
    End If

    For intX = 1 To vas�˻��ڵ�.MaxRows
        If GET_CELL(vas�˻��ڵ�, 1, intX) = "1" Then
            gstrQuy = "INSERT INTO EQ_PROFILE "
            gstrQuy = gstrQuy & vbCrLf & " (EQ_PROFILECD,    equipno,  equipcode,   examcode,   EQ_PROFILENM) "
            gstrQuy = gstrQuy & vbCrLf & " VALUES "
            gstrQuy = gstrQuy & vbCrLf & " ('" & Trim(txtEQ_PROFILECD) & "', "
            gstrQuy = gstrQuy & vbCrLf & "  '" & gtypREG_INFO.EQUIPCD & "', "
            gstrQuy = gstrQuy & vbCrLf & "  '" & GET_CELL(vas�˻��ڵ�, 2, intX) & "', "
            gstrQuy = gstrQuy & vbCrLf & "  '" & GET_CELL(vas�˻��ڵ�, 3, intX) & "', "
            gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtEQ_PROFILENM) & "') "
            If RunSQL(gstrQuy) = False Then
                ADC.RollbackTrans
                Call CloseDB
                Call ErrQuery(gstrQuy, 0)
                Exit Sub
            End If
        End If
    Next intX
    
    ADC.CommitTrans
    
    Call CloseDB
    
    Call MM_KEY_CLEAR
    
    Call MM_VIEW '/���¸� �ٽ� ���̱� ����...
    
    MsgBox "���� �Ǿ����ϴ�.", vbInformation, "Ȯ��"
End Sub

Private Sub Form_Load()
    Call MM_INITIAL
    Call MM_VIEW
End Sub

Private Sub txtEQ_PROFILECD_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtEQ_PROFILECD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEQ_PROFILECD_LostFocus()
    txtEQ_PROFILECD = UCase(txtEQ_PROFILECD)
End Sub

Private Sub txtEQ_PROFILENM_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtEQ_PROFILENM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub vas��������_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim lngRow  As Long
    
    Call MM_KEY_CLEAR
    
    If Row < 1 Then Exit Sub
    
    txtEQ_PROFILECD = GET_CELL(vas��������, 1, Row)
    txtEQ_PROFILENM = GET_CELL(vas��������, 2, Row)
    
    If OpenDB(gtypREG_INFO.DB_CONSTR_LOCAL) = False Then Exit Sub

    gstrQuy = "SELECT * "
    gstrQuy = gstrQuy & vbCrLf & "  FROM EQ_PROFILE "
    gstrQuy = gstrQuy & vbCrLf & " WHERE EQ_PROFILECD = '" & GET_CELL(vas��������, 1, Row) & "' "
    If ReadSQL(gstrQuy, ADR) = False Then
        Call CloseDB
        Exit Sub
    End If
    
    If Not ADR Is Nothing Then
        Do Until ADR.EOF
            For intX = 1 To vas�˻��ڵ�.MaxRows
                If Trim(ADR!EQUIPCODE & "") = GET_CELL(vas�˻��ڵ�, 2, intX) And _
                   Trim(ADR!ExamCode & "") = GET_CELL(vas�˻��ڵ�, 3, intX) Then
                    Call SET_CELL(vas�˻��ڵ�, 1, intX, "1")
                End If
            Next intX
            
            ADR.MoveNext
        Loop
        ADR.Close: Set ADR = Nothing
        
        cmdDelete.Enabled = True
        
        Call vas�˻��ڵ�.SetActiveCell(1, 1)
    End If
    Call CloseDB
End Sub
