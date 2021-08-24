VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frm����_Set_Equipment 
   BorderStyle     =   1  '���� ����
   Caption         =   "Medical Equipment Info Setting"
   ClientHeight    =   5955
   ClientLeft      =   4785
   ClientTop       =   4500
   ClientWidth     =   10695
   BeginProperty Font 
      Name            =   "����ü"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm����_Set_Equipment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   10695
   Begin FPSpread.vaSpread spr�Ƿ���񸮽�Ʈ 
      Height          =   4755
      Left            =   60
      TabIndex        =   3
      Top             =   600
      Width           =   10575
      _Version        =   393216
      _ExtentX        =   18653
      _ExtentY        =   8387
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   23
      MaxRows         =   10
      OperationMode   =   3
      SelectBlockOptions=   0
      SpreadDesigner  =   "frm����_Set_Equipment.frx":9F8A
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
      Height          =   375
      Left            =   8100
      TabIndex        =   2
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "����(&X)"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9420
      TabIndex        =   1
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   60
      X2              =   10620
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "EMR Interface Medical Equipment List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   180
      TabIndex        =   0
      Top             =   60
      Width           =   3915
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   5  '���� �밢��
      Height          =   495
      Index           =   1
      Left            =   60
      Shape           =   4  '�ձ� �簢��
      Top             =   60
      Width           =   10575
   End
End
Attribute VB_Name = "frm����_Set_Equipment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function MM_CANCEL() As Boolean
    Call MM_KEY_CLEAR
End Function

Public Function MM_DELETE() As Boolean

End Function

Private Sub MM_INITIAL()
    Me.Height = 6435
    Me.Width = 10815
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    Call MM_CANCEL
End Sub

Public Function MM_INPUT() As Boolean

End Function

Private Sub MM_KEY_CLEAR()
    If spr�Ƿ���񸮽�Ʈ.MaxRows > 0 Then spr�Ƿ���񸮽�Ʈ.MaxRows = 0
End Sub

Public Function MM_PRINT() As Boolean

End Function

Public Function MM_SAVE() As Boolean
    MM_SAVE = False
    
    '/��� �Ƿ���� ����(��������) �������� Setting
    gtypEQ_INFO.EQUIPCODE = GET_CELL(spr�Ƿ���񸮽�Ʈ, 1, spr�Ƿ���񸮽�Ʈ.ActiveRow)
    gtypEQ_INFO.EQUIPNM = GET_CELL(spr�Ƿ���񸮽�Ʈ, 2, spr�Ƿ���񸮽�Ʈ.ActiveRow)
    gtypEQ_INFO.EQUIPSEQ = GET_CELL(spr�Ƿ���񸮽�Ʈ, 3, spr�Ƿ���񸮽�Ʈ.ActiveRow)
    gtypEQ_INFO.EQIMGFILEPATH = GET_CELL(spr�Ƿ���񸮽�Ʈ, 21, spr�Ƿ���񸮽�Ʈ.ActiveRow) '/���������������(Client)
    gtypEQ_INFO.FTPIMGFILEPATH = GET_CELL(spr�Ƿ���񸮽�Ʈ, 22, spr�Ƿ���񸮽�Ʈ.ActiveRow) '/FTP�������������(Client)
    
    MM_SAVE = True
End Function

Public Function MM_UPDATE() As Boolean

End Function

Public Function MM_VIEW() As Boolean
    Dim strEQCD             As String
    Dim strEQSEQ            As String
    
    Dim strEQCD_Array
    Dim strEQSEQ_Array
    
    Call MM_KEY_CLEAR
    
    '/��� �Ƿ���� ����(��������) ��������
    strEQCD = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQCD)
    strEQSEQ = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQSEQ)
    
    strEQCD_Array = Split(strEQCD, ",")
    strEQSEQ_Array = Split(strEQSEQ, ",")
    
    If OpenDB(gstrREG_DB_CONSTR) = True Then
        With spr�Ƿ���񸮽�Ʈ
            For intX = 0 To UBound(strEQCD_Array)
                gstrQuy = "SELECT A.*, B.EQUIPNM, B.EQUIPTYPE "
                gstrQuy = gstrQuy & vbCrLf & "  FROM MM_EMR_CONF A INNER JOIN MM_EMR_EQUIP B ON A.EQUIPCODE = B.EQUIPCODE "
                gstrQuy = gstrQuy & vbCrLf & " WHERE A.EQUIPCODE = '" & strEQCD_Array(intX) & "' "
                gstrQuy = gstrQuy & vbCrLf & "   AND A.EQUIPSEQ  =  " & Val(strEQSEQ_Array(intX)) & " "
                If ReadSQL(gstrQuy, ADR) = False Then Call CloseDB: End
                    
                If Not ADR Is Nothing Then
                    .MaxRows = .MaxRows + 1: .Row = .MaxRows
                
                    .Col = 1: .Text = Trim(ADR!EQUIPCODE & "")  '/����ڵ�
                    .Col = 2: .Text = Trim(ADR!EQUIPNM & "")    '/����
                    .Col = 3: .Text = Trim(ADR!EQUIPSEQ & "")   '/���SEQ
                    .Col = 4: .Text = Trim(ADR!DEPTCODE & "")   '/DEPTCODE(��ġ���)
                    .Col = 5:                                   '/��������(1.SM, 2.VPM, 3.ICM)
                    Select Case Trim(ADR!EQUIPTYPE & "")
                        Case "1":   .Text = "1.SM"
                        Case "2":   .Text = "2.VPM"
                        Case "3":   .Text = "3.ICM"
                    End Select
                    .Col = 6:                                   '/Imageȹ����(1.����, 2.����)
                    Select Case Trim(ADR!RECEIVETYPE & "")
                        Case "1":   .Text = "1.����"
                        Case "2":   .Text = "2.����"
                    End Select
                    .Col = 7: .Text = Trim(ADR!EQUIPPORT & "")  '/���PC������Ʈ(RECEIVETYPE �����ϰ��)
                    .Col = 8:                                   '/ó������(Y.ó��, N.��ó��)
                    Select Case Trim(ADR!ORDYN & "")
                        Case "Y":   .Text = "Y.ó��"
                        Case "N":   .Text = "N.��ó��"
                    End Select
                    .Col = 9:                                  '/ó����������(ó�������� Y�� ��:1.������QUARY, 2.ó��QUARY, 3.���հ���)
                    Select Case Trim(ADR!QUERYTYPE & "")
                        Case "1":   .Text = "1.������QUARY"
                        Case "2":   .Text = "2.ó��QUARY"
                        Case "3":   .Text = "3.���հ���"
                    End Select
                    .Col = 10:                                  '/ZanImagePrinter��뿩��(Y.���,N.�̻��)
                    Select Case Trim(ADR!ZIPYN & "")
                        Case "Y":   .Text = "Y.���"
                        Case "N":   .Text = "N.�̻��"
                    End Select
                    .Col = 11: .Text = Trim(ADR!ZIPNM & "")     '/ZanImagePrinter Device Name
                    .Col = 12:                                  '/RS232 SERIAL ��뿩��(Y.���, N.�̻��)
                    Select Case Trim(ADR!SERIALYN & "")
                        Case "Y":   .Text = "Y.���"
                        Case "N":   .Text = "N.�̻��"
                    End Select
                    .Col = 13: .Text = Trim(ADR!SERIALPORT & "")        '/RS232 SERIAL PORT
                    .Col = 14: .Text = Trim(ADR!SERIALBAUD & "")        '/RS232 ��żӵ�
                    .Col = 15: .Text = Trim(ADR!SERIALDATABIT & "")     '/RS232 DATABIT(7,8)
                    .Col = 16: .Text = Trim(ADR!SERIALSTARTBIT & "")    '/RS232 STARTBIT(1,2)
                    .Col = 17: .Text = Trim(ADR!SERIALSTOPBIT & "")     '/RS232 STOPBIT(1,2)
                    .Col = 18: .Text = Trim(ADR!SERIALPARITY & "")      '/RS232 PARITY(E,N,O)
                    .Col = 19: .Text = Trim(ADR!SERIALRTS & "")         '/RS232 RTS(0,1)
                    .Col = 20: .Text = Trim(ADR!SERIALDTR & "")         '/RS232 DTR(0,1)
                    .Col = 21: .Text = Trim(ADR!EQIMGFILEPATH & "")     '/���������������(Client)
                    .Col = 22: .Text = Trim(ADR!FTPIMGFILEPATH & "")    '/FTP�������������(Client)
                    .Col = 23: .Text = Trim(ADR!REMARK & "")            '/���
                    
                    If .MaxTextRowHeight(.Row) > 13.3 Then .RowHeight(.Row) = .MaxTextRowHeight(.Row)
                    
                    ADR.Close: Set ADR = Nothing
                End If
            Next intX
        End With
        
        Call CloseDB
    End If
End Function

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdSave_Click()
    If spr�Ƿ���񸮽�Ʈ.ActiveRow < 1 Then
        MsgBox "������ �Ƿ���� �����ϴ�!", vbCritical, "���úҰ�": Exit Sub
    End If
    If MM_SAVE = True Then Unload Me
End Sub

Private Sub Form_Load()
    Call MM_INITIAL
    
    Call MM_VIEW
End Sub

Private Sub spr�Ƿ���񸮽�Ʈ_DblClick(ByVal Col As Long, ByVal Row As Long)
    If spr�Ƿ���񸮽�Ʈ.Row < 1 Then
        MsgBox "������ �Ƿ���� �����ϴ�!", vbCritical, "���úҰ�": Exit Sub
    End If
    If MM_SAVE = True Then Unload Me
End Sub

Private Sub spr�Ƿ���񸮽�Ʈ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If spr�Ƿ���񸮽�Ʈ.ActiveRow < 1 Then
            MsgBox "������ �Ƿ���� �����ϴ�!", vbCritical, "���úҰ�": Exit Sub
        End If
        If MM_SAVE = True Then Unload Me
    End If
End Sub
