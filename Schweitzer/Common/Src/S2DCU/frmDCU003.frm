VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmDCU003 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "��������"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11535
   HelpContextID   =   41001
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.PictureBox fraMesg 
      BackColor       =   &H00DBE6E6&
      Height          =   5175
      Left            =   3405
      ScaleHeight     =   5115
      ScaleWidth      =   7815
      TabIndex        =   9
      Top             =   1095
      Width           =   7875
      Begin RichTextLib.RichTextBox Rtxt 
         Height          =   4050
         Left            =   75
         TabIndex        =   10
         Top             =   990
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   7144
         _Version        =   393217
         TextRTF         =   $"frmDCU003.frx":0000
      End
      Begin MedControls1.LisLabel lblWriteNm 
         Height          =   315
         Left            =   1005
         TabIndex        =   12
         Top             =   645
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblWriteDt 
         Height          =   330
         Left            =   5565
         TabIndex        =   13
         Top             =   630
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   582
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
      End
      Begin VB.TextBox txtText 
         BackColor       =   &H00EFFEFE&
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4050
         Left            =   75
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '����
         TabIndex        =   11
         Top             =   990
         Width           =   7665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Today's Notice"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00DF6A3E&
         Height          =   225
         Left            =   3045
         TabIndex        =   17
         Top             =   120
         Width           =   1710
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   150
         X2              =   7665
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label lblDate 
         BackStyle       =   0  '����
         Caption         =   "�Խ��� :"
         Height          =   195
         Left            =   4755
         TabIndex        =   16
         Top             =   720
         Width           =   690
      End
      Begin VB.Label lblWriter 
         BackStyle       =   0  '����
         Caption         =   "�۾��� :"
         Height          =   195
         Left            =   195
         TabIndex        =   15
         Top             =   705
         Width           =   690
      End
      Begin VB.Label Label1 
         BackColor       =   &H00F7FFFF&
         BorderStyle     =   1  '���� ����
         Height          =   4050
         Left            =   75
         TabIndex        =   14
         Top             =   990
         Width           =   7665
      End
   End
   Begin Crystal.CrystalReport CReport 
      Left            =   6570
      Top             =   6420
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Caption         =   "���(&P)"
      Enabled         =   0   'False
      Height          =   435
      Left            =   7185
      Style           =   1  '�׷���
      TabIndex        =   8
      Top             =   6405
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.OptionButton optWorkFg 
      BackColor       =   &H00DBE6E6&
      Caption         =   "��������"
      Height          =   255
      Index           =   0
      Left            =   3060
      TabIndex        =   7
      Top             =   540
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.OptionButton optWorkFg 
      BackColor       =   &H00DBE6E6&
      Caption         =   "��������"
      Height          =   255
      Index           =   1
      Left            =   4500
      TabIndex        =   6
      Top             =   540
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "����(&X)"
      Height          =   435
      Left            =   9930
      Style           =   1  '�׷���
      TabIndex        =   3
      Top             =   6405
      Width           =   1305
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00E0E0E0&
      Caption         =   "����(&D)"
      Height          =   435
      Left            =   8557
      Style           =   1  '�׷���
      TabIndex        =   2
      Top             =   6405
      Width           =   1305
   End
   Begin VB.CheckBox chkLoadAtStartup 
      BackColor       =   &H00DBE6E6&
      Caption         =   "���� �� ǥ��(&S)"
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   6435
      Width           =   1755
   End
   Begin FPSpread.vaSpread tblTitle 
      Height          =   5160
      Left            =   195
      TabIndex        =   1
      Top             =   1095
      Width           =   3195
      _Version        =   196608
      _ExtentX        =   5636
      _ExtentY        =   9102
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   1
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      MaxCols         =   5
      MaxRows         =   10
      NoBorder        =   -1  'True
      OperationMode   =   1
      Protect         =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "frmDCU003.frx":009D
      VisibleCols     =   2
      VisibleRows     =   10
      TextTip         =   3
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '��� ����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EE772F&
      Height          =   315
      Left            =   1005
      TabIndex        =   5
      Top             =   450
      Width           =   1350
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   300
      Picture         =   "frmDCU003.frx":0476
      Top             =   360
      Width           =   630
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  '�ܻ�
      Height          =   630
      Index           =   0
      Left            =   180
      Shape           =   4  '�ձ� �簢��
      Top             =   300
      Width           =   2355
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      FillColor       =   &H00404040&
      FillStyle       =   0  '�ܻ�
      Height          =   630
      Index           =   1
      Left            =   240
      Shape           =   4  '�ձ� �簢��
      Top             =   360
      Width           =   2355
   End
   Begin VB.Label lblToday 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "1999�� 11�� 12�� �ݿ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007A5268&
      Height          =   240
      Left            =   8700
      TabIndex        =   4
      Top             =   780
      Width           =   2580
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00808080&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  '�ܻ�
      Height          =   315
      Index           =   2
      Left            =   8700
      Shape           =   4  '�ձ� �簢��
      Top             =   720
      Width           =   2580
   End
End
Attribute VB_Name = "frmDCU003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+--------------------------------------------------------------------------------------+
'|  1. Form��   : frmDCU003
'|  2. ��  ��   : �������� List
'|  3. �ۼ���   : �� ����
'|  4. �ۼ���   : 2000.11.6
'|
'|  CopyRight(C) 2002 Pomis
'+--------------------------------------------------------------------------------------+
Option Explicit

Private ObjSql As clsDCUSqlStmt
Private mvarProjectId As String 'APS, BBS, LIS ���θ� �޾ƿ��� ����
Private mvarTradeMark As String '
Private mvarCanDelete As Boolean
Private strPID As String        'APS, BBS, LIS �޴� ����
Dim strEntDt As String          '�Է³�¥ ���� ����
Dim lngSeq As Long              '�Ϸù�ȣ ���� ����

Public Property Let ProjectId(ByVal vData As String)
    mvarProjectId = vData
End Property

Public Property Get ProjectId() As String
    ProjectId = mvarProjectId
End Property

Public Property Let TradeMark(ByVal vData As String)
    mvarTradeMark = vData
End Property

Public Property Get TradeMark() As String
    TradeMark = mvarTradeMark
End Property

Public Property Let CanDelete(ByVal vData As Boolean)
    mvarCanDelete = vData
End Property

Public Property Get CanDelete() As Boolean
    CanDelete = mvarCanDelete
End Property

Private Sub cmdDelete_Click()
    Dim strTmp As VbMsgBoxResult
    
    '�������� Ȯ��...
    strTmp = MsgBox("�����Ͻðڽ��ϱ�?", vbInformation + vbOKCancel, Me.Caption)
    If strTmp = vbCancel Then
        Exit Sub
    Else
        '����
        Set ObjSql = New clsDCUSqlStmt
        If ObjSql.Del_COM011(lngSeq, strEntDt) = True Then
                MsgBox "�����Ͽ����ϴ�.", vbInformation, Me.Caption
        End If
        Set ObjSql = Nothing
    End If
    'Spread�� load ��Ű��..
    LoadTitles
    Label2.Caption = "Today's Notice"
    lblWriteNm.Caption = ""
    lblWriteDt.Caption = ""
    txtText = ""
    Rtxt.TextRTF = ""
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'-------------------------------'
'   2002-12-09 �ۼ��� : �̻��
'-------------------------------'
Private Sub cmdPrint_Click()
    Dim strRptPath  As String   '����Ʈ���� ���
    Dim strHospNm   As String   '������
    Dim SQL         As String
    
    
    strEntDt = Format(lblWriteDt.Caption, "yyyymmdd")
    
On Error GoTo ErrPrint
    If optWorkFg(0).Value = True Then
        strRptPath = InstallDir & "LIS\Rpt\rptSubReport.rpt"
        strHospNm = medGetP(medGetINI("LIS_CONST", "P_HOSPITALNAME", INIPath), 1, LINE_DIV)
        
        With CReport
            .Connect = "DSN=" & ObjSysInfo.DatabaseNm & ";UID=" & ObjSysInfo.DBLoginId & ";PWD=" & ObjSysInfo.DBPassword & ";"
            .SQLQuery = "SELECT users, note FROM " & T_COM011 & _
                                 " WHERE seq=" & lngSeq & " AND inputday='" & strEntDt & "'"
            .ParameterFields(0) = "Depart;" & Label2.Caption & ";true"
            .ParameterFields(1) = "Name;" & lblWriteNm.Caption & ";true"
            .ParameterFields(2) = "Date;" & lblWriteDt.Caption & ";true"
            .ParameterFields(3) = "Date;" & lblWriteDt.Caption & ";true"
            .ParameterFields(4) = "HosptalNm;" & strHospNm & " ���ܰ˻����а�" & ";true"
            .ParameterFields(5) = "Name1;" & " " & ";true"
            .ParameterFields(6) = "Name2;" & " " & ";true"
            .ReportFileName = strRptPath
            .WindowShowRefreshBtn = True
            .RetrieveDataFiles
            .WindowState = crptMaximized
            .Destination = crptToWindow
            .Action = 1
            .Reset
        End With
    Else
        strRptPath = InstallDir & "LIS\Rpt\rptDailyReport.rpt"
        strHospNm = medGetP(medGetINI("LIS_CONST", "P_HOSPITALNAME", INIPath), 1, LINE_DIV)
        
        With CReport
            .Connect = "DSN=" & ObjSysInfo.DatabaseNm & ";UID=" & ObjSysInfo.DBLoginId & ";PWD=" & ObjSysInfo.DBPassword & ";"
            .SQLQuery = "SELECT users, note FROM " & T_COM011 & _
                                 " WHERE seq=" & lngSeq & " AND inputday='" & strEntDt & "'"
            .ParameterFields(0) = "Depart;" & Label2.Caption & ";true"
            .ParameterFields(1) = "Name;" & lblWriteNm.Caption & ";true"
            .ParameterFields(2) = "Date;" & lblWriteDt.Caption & ";true"
            .ParameterFields(3) = "HosptalNm;" & strHospNm & " ���ܰ˻����а�" & ";true"
            .ParameterFields(4) = "Name1;" & " " & ";true"
            .ParameterFields(5) = "Name2;" & " " & ";true"
            .ReportFileName = strRptPath
            .WindowShowRefreshBtn = True
            .RetrieveDataFiles
            .WindowState = crptMaximized
            .Destination = crptToWindow
            .Action = 1
            .Reset
        End With
    End If
    
    Exit Sub
        
ErrPrint:
    MsgBox Err.Description, vbCritical, "����"
End Sub

Private Sub Form_Activate()
    cmdDelete.Enabled = mvarCanDelete
End Sub

'-------------------------------'
'   2002-12-09 ������ : �̻��
'-------------------------------'
Private Sub Form_Load()
    Dim strAppName As String
    
    '���� ��¥�� ���� ����..
    lblToday.Caption = Format(GetSystemDate, "dddddd")

    Rtxt.Visible = False: txtText.Visible = True: txtText.ZOrder 0

    
    Call LoadTitles

    strAppName = mvarTradeMark & " " & mvarProjectId
    chkLoadAtStartup.Value = GetSetting(strAppName, "Options", "ShowAtStart", 0)

    '��������, �������� ǥ��
    If ObjMyUser.IsManager Or ObjMyUser.IsDeveloper Or ObjMyUser.IsSupervisor Then
        optWorkFg(0).Visible = True
        optWorkFg(1).Visible = True
        cmdPrint.Visible = True
    End If
End Sub

'-------------------------------'
'   2002-08-06 ������ : �̻��
'-------------------------------'
Private Sub LoadTitles()
    Dim Rs As Recordset
    Dim SqlStmt As String
    Dim i As Integer
    
    '������ ��������.
    Set ObjSql = New clsDCUSqlStmt

    If optWorkFg(0).Value = True Then
        ObjSql.WorkFg = "0"
    ElseIf optWorkFg(1).Value = True Then
        ObjSql.WorkFg = "1"
    End If
    
    Set Rs = New Recordset
    
    Rs.Open ObjSql.GetSQLCOM011ByDeptFg2(mvarProjectId), DBConn
    
    '����Ÿ�� ������ �ѱ���...
    If Rs.EOF Then
        With tblTitle
            .Col = -1
            .Row = -1
            .Text = ""
        End With

        txtText = "�������� ������ �����ϴ�."
        
        cmdPrint.Enabled = False
        
        Set Rs = Nothing
        Set ObjSql = Nothing
        Exit Sub
    Else
        cmdPrint.Enabled = True
    End If
    
    'Spread�� load .....
    With tblTitle
        .MaxRows = Rs.RecordCount
        For i = 1 To Rs.RecordCount
            .Row = i
            .Col = 1: .Text = Format("" & Format(Rs.Fields("inputDay").Value, CS_DateMask), "YY-MM-DD")
            .Col = 2: .Text = "" & Rs.Fields("Title").Value
            .Col = 3: .Text = "" & Rs.Fields("inputDay").Value
            .Col = 4: .Text = Trim("" & Rs.Fields("seq").Value)
            .Col = 5: .Text = "" & Rs.Fields("users").Value
            Rs.MoveNext
        Next
        If .MaxRows > 0 Then Call tblTitle_Click(1, 1)
    End With
    Set Rs = Nothing
    Set ObjSql = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strAppName As String
    
    strAppName = mvarTradeMark & " " & mvarProjectId
    SaveSetting strAppName, "Options", "ShowAtStart", chkLoadAtStartup.Value

End Sub

'-------------------------------'
'   2002-08-06 �ۼ��� : �̻��
'-------------------------------'
Private Sub optWorkFg_Click(Index As Integer)
    If Index = 0 Then
        frmDCU003.Caption = "��������"
        lblTitle.Caption = "��������"
        cmdPrint.Enabled = True
        Call LoadTitles
    ElseIf Index = 1 Then
        frmDCU003.Caption = "��������"
        lblTitle.Caption = "��������"
        cmdPrint.Enabled = True
        Call LoadTitles
    End If
End Sub

Private Sub tblTitle_Click(ByVal Col As Long, ByVal Row As Long)
    Dim Rs As Recordset
    
    If tblTitle.DataRowCnt < 1 Then Exit Sub
    
    '���̺��� Ŭ���ϸ� ���� �ѱ���....
    With tblTitle
        .Col = -1
        .Row = -1
        .ForeColor = &H404040
        .Row = Row
        .ForeColor = &HFF0000
        .Col = 1: 'lblWriteDt.Caption = .Text
        .Col = 2: Label2.Caption = .Text
        .Col = 3: strEntDt = .Text: lblWriteDt.Caption = Format(.Text, CS_DateMask)
        .Col = 4: lngSeq = (.Text)
        .Col = 5: lblWriteNm.Caption = .Text
    End With
    '������ �ִ��� üũ����...
    Set ObjSql = New clsDCUSqlStmt
    
    Set Rs = New Recordset
    Rs.Open ObjSql.GetSQLCOM011BySeq(lngSeq, strEntDt), DBConn
    
    '������ ������ �ѱ���...
    If Rs.EOF Then
        Set Rs = Nothing
        Set ObjSql = Nothing
        Exit Sub
    End If
    
    txtText.Text = "" & Rs.Fields("note").Value

    Set Rs = Nothing
    Set ObjSql = Nothing
End Sub
