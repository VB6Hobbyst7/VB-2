VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmVitros 
   Caption         =   "Vitros ECI"
   ClientHeight    =   8490
   ClientLeft      =   135
   ClientTop       =   645
   ClientWidth     =   11880
   Icon            =   "frmVitros.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  '�ִ�ȭ
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  '�� ����
      Height          =   660
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "�ڷ������ �����մϴ�"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "�ڷ������ �����մϴ�"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   ""
            Object.ToolTipText     =   "ȭ���� ����Ÿ�� DB�� �����մϴ�"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "ȭ���� ����Ÿ�� ����ϴ�"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "���ȯ���� �����մϴ� "
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "���α׷��� �����մϴ�"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.PictureBox picResult 
      Height          =   6060
      Left            =   6030
      ScaleHeight     =   6000
      ScaleWidth      =   5310
      TabIndex        =   13
      Top             =   1575
      Visible         =   0   'False
      Width           =   5370
      Begin FPSpread.vaSpread SSR 
         Height          =   5820
         Left            =   135
         TabIndex        =   14
         Top             =   90
         Width           =   5055
         _Version        =   196608
         _ExtentX        =   8916
         _ExtentY        =   10266
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   3
         MaxRows         =   50
         SpreadDesigner  =   "frmVitros.frx":030A
      End
   End
   Begin FPSpread.vaSpread SS 
      Height          =   5250
      Left            =   225
      TabIndex        =   12
      Top             =   1800
      Width           =   11175
      _Version        =   196608
      _ExtentX        =   19711
      _ExtentY        =   9260
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   6
      EditEnterAction =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   30
      SpreadDesigner  =   "frmVitros.frx":0EF9
      UserResize      =   1
      VisibleCols     =   23
      VisibleRows     =   120
   End
   Begin VB.TextBox txtBarCode 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   16
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Left            =   8976
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2184
      Visible         =   0   'False
      Width           =   636
   End
   Begin VB.FileListBox File1 
      Height          =   270
      Left            =   9612
      TabIndex        =   5
      Top             =   2184
      Visible         =   0   'False
      Width           =   996
   End
   Begin VB.Timer Timer_RRequest 
      Left            =   10320
      Top             =   1728
   End
   Begin VB.Timer Timer_RCheck 
      Left            =   9984
      Top             =   1728
   End
   Begin VB.Timer Timer_Picture 
      Left            =   9312
      Top             =   1728
   End
   Begin VB.Timer Timer1 
      Left            =   8976
      Top             =   1728
   End
   Begin VB.Timer Timer_Order 
      Left            =   9648
      Top             =   1728
   End
   Begin VB.ListBox ErrList 
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   4968
      TabIndex        =   3
      Top             =   7248
      Width           =   6444
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3312
      TabIndex        =   2
      Top             =   2355
      Width           =   2868
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   432
      TabIndex        =   1
      Top             =   2355
      Width           =   2868
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10608
      Top             =   2184
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   11052
      Top             =   1728
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Frame Frame2 
      Height          =   1308
      Left            =   288
      TabIndex        =   8
      Top             =   7152
      Width           =   4524
      Begin Threed.SSPanel SSPan 
         Height          =   492
         Left            =   144
         TabIndex        =   9
         Top             =   168
         Width           =   4236
         _Version        =   65536
         _ExtentX        =   7472
         _ExtentY        =   868
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   8.95
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelOuter      =   1
         BevelInner      =   2
      End
      Begin VB.Label lblTime 
         Alignment       =   2  '��� ����
         BorderStyle     =   1  '���� ����
         ForeColor       =   &H00FF0000&
         Height          =   492
         Left            =   2304
         TabIndex        =   11
         Top             =   720
         Width           =   2076
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  '���� ����
         Height          =   492
         Left            =   144
         TabIndex        =   10
         Top             =   720
         Width           =   2076
      End
   End
   Begin VB.Label lblPort 
      Alignment       =   2  '��� ����
      Height          =   285
      Left            =   4980
      TabIndex        =   17
      Top             =   8100
      Width           =   1965
   End
   Begin VB.Label Label3 
      Caption         =   "BarCode�Է�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1530
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   5
      Left            =   7560
      Picture         =   "frmVitros.frx":2F58
      Top             =   930
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   4
      Left            =   7170
      Picture         =   "frmVitros.frx":3262
      Top             =   930
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   6795
      Picture         =   "frmVitros.frx":356C
      Top             =   930
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   6405
      Picture         =   "frmVitros.frx":3876
      Top             =   930
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   6030
      Picture         =   "frmVitros.frx":3B80
      Top             =   930
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   5640
      Picture         =   "frmVitros.frx":3E8A
      Top             =   930
      Visible         =   0   'False
      Width           =   480
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   11004
      Top             =   2184
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmVitros.frx":4194
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmVitros.frx":44AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmVitros.frx":47C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmVitros.frx":4AE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmVitros.frx":4DFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmVitros.frx":5116
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "date ����� obj"
      Height          =   228
      Left            =   8976
      TabIndex        =   4
      Top             =   2712
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   492
      Left            =   864
      TabIndex        =   0
      Top             =   840
      Width           =   4932
   End
   Begin VB.Menu mnuOption 
      Caption         =   "�ɼ�(&M)"
      Visible         =   0   'False
      Begin VB.Menu mnuRack 
         Caption         =   "Rack/Position Set"
      End
   End
   Begin VB.Menu mnuReceive 
      Caption         =   "�ڷ����(&R)"
   End
   Begin VB.Menu mnuEnd 
      Caption         =   "��������(&E)"
   End
   Begin VB.Menu mnuWrite 
      Caption         =   "�ڷ�����(&W)"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuClear 
      Caption         =   "ȭ�������(&C)"
   End
   Begin VB.Menu mnuSet 
      Caption         =   "ȯ�漳��(&S)"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "����(&X)"
   End
End
Attribute VB_Name = "frmVitros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Dim RPoint                                  'row ��ġ ������
    Dim CPoint                                  'col ��ġ ������
    Dim RSequence                               'Result Record Sequence Check��
    
    Dim RResult
    
    Dim Update_Check        As Boolean          ' ���� check��
    Dim Update_Check_Force  As Boolean          ' ���� check��
    Dim Receive_Check       As Boolean
    
    Dim GBTransmit          As String
    Dim strBiDirect_Trans   As Boolean          ' batch�� data ���Ž� ���
    
    Dim Tcounter                                ' ����ǥ�� image count��
    Dim QCounter            As Integer
    Dim RCounter            As Integer          ' data �۽Ž� record count check ��
    
    Dim Ser                 As Integer
    Dim ResultText          As String
    Dim i
    Dim j
    
    Dim C1                  As Integer          ' work buffer column1
    Dim R1                  As Integer          ' work buffer row1
    
    Dim Temp_Jeobsu(100, 50)            As String           ' Spread�� Data�� Temp_Jeobsu Array�� Move
    Dim Temp_Request_Code(100, 50)      As String           ' input test code A1~Z6
    Dim Temp_Result(100, 50)            As String           ' Result Input Data
    
    Dim Temp_JCount                     As String           ' Spread�� Data�� Temp_Jeobsu Array�� Move
    Dim Temp_Quality(10, 100)           As String
    
    Dim Temp_K(60, 3)       As String           ' item table data �Է¿� buffer
    Dim FileClose           As Boolean
    
    Dim LNormal             As Boolean          'working list msg termination flag
    
    Dim N                   As Integer
    Dim JeobsuCheck         As Boolean          'Data_Update
    Dim Pflag               As Boolean          'Data_Update
    
    Dim RecordCountSum                          'Data_Update
    Dim RecordCountBit                          'Data_Update
    
    Dim MaxRecordCount      As Long             '�˻� code count
    
    Dim MaxDataRowCnt       As Integer
    
    Dim SColumn
    Dim SRow
    
    Dim temp_file           As String           'file directory
    Dim temp_file_Order     As String           'file directory
    
    Dim hSaveFile
    Dim hSaveFile_Order
    
    Dim RBuffer             As String
    Dim RBufferSum          As String
    Dim RCheckSum           As String
    Dim RCheckSumD          As String
    
    
    Dim SBuffer             As String
    Dim SBufferSum          As String
    Dim SBufferSumD         As String
    
    Dim SOHBuffer           As Boolean
    Dim STXBuffer           As Boolean
    Dim ETXBuffer           As Boolean
    Dim EOTBuffer           As Boolean
    Dim ENQBuffer           As Boolean
    Dim ACKBuffer           As Boolean
    Dim NACKBuffer          As Boolean
    
    Dim SOH                 As String
    Dim STX                 As String
    Dim ETX                 As String
    Dim EOT                 As String
    Dim ENQ                 As String
    Dim ACK                 As String
    Dim NACK                As String
    Dim ETB                 As String
    
    Dim Order_Data_Seq      As Integer
    Dim Or_Seq              As Integer
    
    Dim timerx              As Boolean
    Dim PortOpen            As Boolean
    Dim SSCheck             As Boolean
    
    Dim BC                  As String           ' Block Code
    Dim LC                  As Integer          ' Data Line Line Code
    Dim DataLine            As String           ' Data Line Type
    
'    Dim Receive_STX_Check   As Boolean
'    Dim Receive_Data_Check  As Boolean
    Dim Receive_STA_Check   As Boolean
    Dim Receive_STA_Seq     As Integer
    
    Dim RackNo_Result       As String
    Dim PosiNo_Result       As String
    Dim Infostr_Result      As String
    
'    Dim LCpos               As Integer
    
    Dim MaxRackNo
    Dim MaxPosiNo
    Dim End_check           As Boolean
    Dim TimerRNo
    Dim TimerPNo
    Dim TimerTCode
    
    Dim Test_Code           As String           'line code 12�� test code  check
    
    Dim Error_Message       As String
    Dim Error_Message_Block As String
    
    Dim Sample_Result(5)    As String
    Dim RecordCount

    Dim SendTime
    Dim SendBuffW           As String
    Dim SendBuffT           As String

    Dim R_Check             As Boolean
    Dim ME_Check            As Boolean
    Dim MA_Check            As Boolean
    
    'data receive part ����
    Dim Mheader
    Dim Mdate
    
    Dim Lterminator
    
    Dim Fcheck              As Integer
    
    Dim RJeobsuNo           As String
    
    Dim Pns                 As Integer      'Patient information length set Start
    Dim Pne                 As Integer      'Patient information length set Start
    Dim Pinfo1                              'Patient information 16Byte
    Dim Pinfo2                              'Patient information 12Byte
    Dim Pinfo3                              'Patient information  6Byte
    Dim Pinfo4                              'Patient information  4Byte
    
    Dim Optno               As String       '���� Data���� ptno�� check�Ͽ� move
    Dim StrGBER             As String       '���ޱ��� Check

    Dim RJeobsuNo_Error     As Boolean
    


'
'
Public Sub GotoSpreadSet()
    SS.SetFocus                             'clear�� cell active ���·� ����
    SS.Row = 1
    SS.Col = 1
    SS.Action = SS_ACTION_GOTO_CELL
    SS.Action = SS_ACTION_ACTIVE_CELL
    
End Sub



Private Sub picResult_Click()
        picResult.Visible = False

End Sub

Private Sub SSR_Click(ByVal Col As Long, ByVal Row As Long)
        picResult.Visible = False

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    On Error Resume Next
    Select Case Button.Index
            Case 1
                   Call mnuReceive_Click
            Case 2
                   Call mnuEnd_Click
            Case 3
                   Call mnuWrite_Click
            Case 4
                   Call mnuClear_Click
            Case 5
                   Call mnuSet_Click
            Case 6
                   Call mnuExit_Click
    End Select
    
End Sub


Private Sub Form_Initialize()
    Dim Title$
    Dim i           As Integer
    ' ������ �̹� Window�� �ش� Program�� Loading �Ǿ��� ���
    '        Loading �Ǿ��ִ� Program�� Activate �ǵ��� �ϴ� Routine
    '        ���� Loading �Ϸ��� Program �� End ��Ų��
    If App.PrevInstance Then
        Title$ = App.Title
        App.Title = "Temp"
        AppActivate Title$
        End
    End If

End Sub


Private Sub Form_Load()
    ' db connect �ʱ� �۾�
    
    DoEvents
    Me.Show
    
    Dim Title$
    
    DoEvents
    Me.Show
    
    Call DbAdoConnect("TW_MIS_EXAM", "HOSPITAL", "kuh2")
    
    SSPan.Caption = "Server ��ǻ�Ϳ� ���ӵǾ����ϴ�."
    SSPan.ForeColor = Val("&H000000FF&")
    
    Call Parmini                            ' spread �ʱ�ȭ �۾�
    Call vaSpread_Clear(SS, 1, 1, 0, 0)
    Call GetIniFile
    
    Call CodeKy_Search                      ' codeky Read from twexam_itemml
    
    Label1.FontSize = 24
    Label1.BorderStyle = 0
    Label1.Caption = " Vitros ECi "
    
    For i = 7 To SS.MaxCols                 ' spread header�� ��� code�� �ʱ�ȭ
        SS.Row = 0
        SS.Col = i
        If Temp_K(i - 6, 1) <> "" Then
            SS.Text = Temp_K(i - 6, 1)
        Else
            SS.Text = "_"
        End If
    Next i
    
    For i = 1 To SS.MaxRows
        For j = 1 To 6
            SS.Row = i
            SS.Col = j
            SS.Lock = True
        Next j
    Next i
    
'    For i = 1 To MaxRecordCount                 ' spread header�� ��� code�� �ʱ�ȭ
'        SSR.Col = 1
'        SSR.Row = i
'        SSR.Text = "   " & Temp_K(i, 0)
'
'        SSR.Col = 2
'        SSR.Row = i
'        SSR.Text = "   " & Temp_K(i, 1)
'    Next i
'    SSR.MaxRows = MaxRecordCount
    
    Call Kdelete                                '30�� ����� ���� file ����
            
    txtBarCode.SetFocus

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Update_Check_Force = True Then Exit Sub
    If Update_Check = True Then
        Cancel = 1                          ' cancel = 0 (true) �� ��츸 �����
    End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
    If MSComm1.PortOpen = True Then
       MSComm1.PortOpen = False
    End If
    
    Call DbAdoDisConnect
    
    End
    
End Sub


Private Sub ErrCancel_Click()
'�ڷ�������
    For i = 1 To SS.MaxRows
        For j = 1 To 5
            SS.Row = i
            SS.Col = j
            SS.Lock = False
        Next j
    Next i
'    SS.Enabled = True
    
    Timer_Picture.Interval = 0                  'Timer_Picture_Timer End
    Timer_Order.Interval = 0                  'Timer_order_Timer End
    Timer_RCheck.Interval = 0                   'Timer_RCheck_Timer End
    Timer_RRequest.Interval = 0                 'Timer_RRequest_Timer End
    timerx = False
    Image1(0).Visible = False
    Image1(1).Visible = False
    Image1(2).Visible = False
    Image1(3).Visible = False
    Image1(4).Visible = False
    
    Call WorkDisplay(0)
    Close #hSaveFile
    FileClose = True

End Sub


Private Sub SS_Click(ByVal Col As Long, ByVal Row As Long)
    If Col = 3 Or Col = 4 Then
        picResult.Visible = True
       i = 0
       j = 0
       
       SSR.MaxRows = 0
       
       SS.Col = 5
       SS.Row = Row
       
       SSR.MaxRows = Val(SS.Text)
       
       'Temp_Jeobsu(RPoint, CPoint)
'        For i = 1 To Val(SS.Text) 'MaxRecordCount
        For i = 1 To MaxRecordCount
            If Temp_Jeobsu(Row, i + 10) <> "" Then
                j = j + 1
                
                SSR.Col = 1
                SSR.Row = j
                SSR.Text = "  " & Temp_K(i, 0)
                
                SSR.Col = 2
                SSR.Row = j
                SSR.Text = "  " & Temp_K(i, 1)
                
                
                SSR.Col = 3
                SSR.Row = j
                SSR.Text = Temp_Result(Row, i + 10)
            End If
        Next i
'        SSR.MaxRows = MaxRecordCount
        SSR.MaxRows = SSR.DataRowCnt
    Else
        SSR.Col = 1:    SSR.Col2 = SSR.MaxCols
        SSR.Row = 1:    SSR.Row2 = SSR.DataRowCnt
         
        SSR.BlockMode = True
        SSR.Action = SS_ACTION_CLEAR_TEXT
        SSR.BlockMode = False
        picResult.Visible = False
    End If
    
    
End Sub

Private Sub SS_DblClick(ByVal Col As Long, ByVal Row As Long)
'    Dim SSDatex             As String
'
'    If Col <> 1 Or Row > SS.DataRowCnt Then
'        For i = 1 To 100
'            Beep
'        Next i
'        Exit Sub
'    End If
'
'    SS.Col = Col
'    SS.Row = Row

End Sub


Private Sub SS_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    
'    Dim Rs                  As ADODB.Recordset
'
'    Dim strPt               As String
'    Dim Bjeobsudt           As String
'    Dim Bslipno1            As String
'    Dim Bslipno2            As String
'
'    Dim Checkdouble         As String
'    Dim Ptnolen
'    Dim Temp_ptno
'    Dim Temp_name
'
'    SS.Row = Row
'    SS.Col = Col
'
'    If Col >= 2 Then
'        SS.Text = ""
'        SS.SetFocus                             'clear�� cell active ���·� ����
'        Exit Sub
'    End If
'
'    strPt = SS.Text
'
'    If Row = 500 Then
'        MsgBox " Max Sequence Number Reached." & _
'               " ���α׷��� ������Ͽ� �ֽñ� �ٶ��ϴ�. "
'        Exit Sub
'    End If
'
'    If Len(strPt) = 12 Then
'        Bjeobsudt = Mid(strPt, 1, 5)
'        Bslipno1 = Mid(strPt, 6, 2)
'        Bslipno2 = Mid(strPt, 8, 5)
'
'        If Col = 1 And strPt <> "" Then
'            SS.Row = Row - 1
'            If Bslipno2 = SS.Text Then
'
'                SS.Row = Row
'                SS.Text = ""
'                SS.SetFocus                             'clear�� cell active ���·� ����
'                SS.Action = SS_ACTION_ACTIVE_CELL
'                SS.BackColor = &HFF&
'
'                Exit Sub
'            End If
'        End If
'
'        strSQL = ""
'        strSQL = strSQL & " SELECT PTNO "
'        strSQL = strSQL & "   FROM TWEXAM_GENERAL_SUB "                  ' �� MASTER
'        strSQL = strSQL & "  WHERE JEOBSUDT = TO_DATE('" & Bjeobsudt & "','YYYY-MM-DD')"
'        strSQL = strSQL & "    AND SLIPNO1 =   '" & Bslipno1 & "'"        ' �Ϸù�ȣ
'        strSQL = strSQL & "    AND SLIPNO2 =   '" & Bslipno2 & "'"        ' �Ϸù�ȣ
'
'        Result = AdoOpenSet(Rs, strSQL)
'
'        If Result Then
'            Rs.MoveFirst
'            Do While Not Rs.EOF
'                Temp_ptno = Trim$(Rs.Fields("ptno")) & ""
'                Rs.MoveNext
'            Loop
'        Else
'
'            SS.Row = Row
'            SS.Text = ""
'            SS.SetFocus                             'clear�� cell active ���·� ����
'
'            MsgBox " DATABASE�� ��ϵ� ������ ���ų� ������ �߸��Ǿ����ϴ�." & vbCrLf & vbCrLf & _
'                   " DATA�� ���Է� �Ͻʽÿ�." & vbCrLf & vbCrLf & _
'                   " ���Է� �Ŀ��� ���� ERROR�� �߻��� ��� ����Ƿ� ���� �ٶ��ϴ�."
'            SS.Action = SS_ACTION_ACTIVE_CELL
'            Exit Sub
'        End If
'
'        AdoCloseSet Rs
'
'
'        strSQL = ""
'        strSQL = strSQL & " SELECT SNAME "
'        strSQL = strSQL & "   FROM TWBAS_PATIENT "                  ' �� MASTER
'        strSQL = strSQL & "  WHERE PTNO = '" & Temp_ptno & "' "     ' PATIENT NO
'
'        If AdoOpenSet(Rs, strSQL) Then
'            Rs.MoveFirst
'            Do While Not Rs.EOF
'                Temp_name = Trim$(Rs.Fields("sname")) & ""
'                Rs.MoveNext
'            Loop
'        Else
'
'            SS.Row = Row
'            SS.Text = ""
'            SS.SetFocus                             'clear�� cell active ���·� ����
'            SS.Action = SS_ACTION_ACTIVE_CELL
'            MsgBox "DATABASE�� ��ϵ� �̸��� ���ų� ������ �߸��Ǿ����ϴ�." & vbCrLf & _
'                   "PTNO�� ���Է� �Ͻʽÿ�." & vbCrLf & vbCrLf & _
'                   "���Է� �Ŀ��� ��� ���� ERROR�� �߻��� ��� ����Ƿ� ���� �ٶ��ϴ�."
'            Exit Sub
'        End If
'
'
'        SS.Row = Row
'
'        SS.Col = 2
'        SS.Text = Temp_ptno
'
'        SS.Col = 3
'        SS.Text = Temp_name
'
'    End If

End Sub

'�̻��
Private Sub SS_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim SS_M_col            As Long             ' spread mouse down col val
    Dim SS_M_row            As Long             ' spread mouse down row val
    
    Dim Msg, Style, Title, Response
    
    Exit Sub
    
    If Receive_Check = True Then
        Exit Sub
    End If
    
    Call SS.GetCellFromScreenCoord(SS_M_col, SS_M_row, x, y)
    
    If (SS_M_col <> 1 Or SS_M_row > SS.DataRowCnt) And Button = vbRightButton Then
        For i = 1 To 50
            Beep
        Next i
        Exit Sub
    End If
    
    If Update_Check = True Then Exit Sub
    
    If Button = vbRightButton Then                                      ' Value = 2
'        Call SS.GetCellFromScreenCoord(SS_M_col, SS_M_row, x, y)
        SS.Col = SS_M_col
        SS.Row = SS_M_row
        If SS.ActiveRow = SS_M_row And SS.ActiveCol = 1 And SS_M_row <= SS.DataRowCnt Then
            Msg = SS_M_row & " ��° DATA�� ���� �Ͻðڽ��ϱ�?" & vbCrLf & _
                             " DATA�� Ȯ���ϼ̽��ϱ�?"
            Style = vbYesNo + vbQuestion + vbDefaultButton2             ' Define buttons.
            Title = "DATA ����"                                         ' �⺻ ����.
            Response = MsgBox(Msg, Style, Title)
            If Response = vbYes Then                                    ' ����ڰ� ���� ����.
                SS.Action = SS_ACTION_DELETE_ROW                        ' value = 5
                
                'Work ���� ���� �۾�
                
            End If
        End If
    End If
    
End Sub



'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Private Sub mnuReceive_Click()
'1)�˻����
    On Error Resume Next

    Receive_Check = True
    
    If SS.DataRowCnt = 0 Then       ' Transmition of Patient File to STA
        strBiDirect_Trans = False   ' batch mode
    Else
        strBiDirect_Trans = True    ' Query Mode
    End If
    
    QCounter = 0
    RCounter = 1

    ErrList.Clear
    
    R_Check = False
    ME_Check = False
    MA_Check = False
    
    timerx = True                                           '����ǥ�� ������ flag
    FileClose = False
    Timer_Picture.Interval = 500                            'Timer_Picture_Timer 1000mS
    N = 0
    
    SColumn = 6                                             'spread sheet �ʱ�ȭ ��ġ
    SRow = 1
    Call SS_INIT(SS, SColumn, SRow)

'******************** Part 3 **********************************
'*    Output�� File Name Set Open/Save ó��                   *
'**************************************************************
    
    On Error GoTo ErrorMsg
    Ser = Ser + 1
    CommonDialog1.InitDir = "C:\intdown"
    CommonDialog1.FileName = "V" & Format$(lblDate, "yyyymmdd") & Ser
    
    'CommonDialog1.Flags = "&H2"                             '�����ߺ� check & msgbox
    CommonDialog1.Filter = "All Files (*.*)|*.*|" & _
                           "Text Files" & "(*.txt)|*.txt|" & _
                           "Ifc Files" & "(*.ifc)|*.ifc|"
    CommonDialog1.FilterIndex = 3
    CommonDialog1.CancelError = True                        '���Ȯ���� ���� ���
    
    On Error GoTo ErrCancel
    CommonDialog1.ShowSave                                  'dialog show
    CommonDialog1.CancelError = True                        'cancel error reset
    On Error GoTo ErrorMsg
    
    DoEvents
    Me.Show
    
    temp_file = CommonDialog1.FileName
    Call WorkDisplay(1)                     '  "���Ϸ� ���� ���Դϴ�." MsgBox Display
    
    If MSComm1.PortOpen = False Then
        MSComm1.InBufferSize = 8192         ' InBufferSize ������ portopen = false�� ��츸 ����  ' default = 1024
        MSComm1.PortOpen = True
    End If
    
    MSComm1.RThreshold = 1                  ' MSComm ��Ʈ���� ���� ���ۿ� �� ���ڰ� ��� �� ������ OnComm �̺�Ʈ�� �߻���ŵ�ϴ�.
    MSComm1.InputLen = 1
    
'******************** Part 4 **********************************
'*  Output File OPen                                          *
'**************************************************************
    hSaveFile = FreeFile
    Open temp_file For Append As hSaveFile
        If Err Then
            MsgBox Error$, vbExclamation
            Close hSaveFile
            hSaveFile = 0
            Call WorkDisplay(0)
            Exit Sub
        End If
    
    Timer_Order.Interval = 30000           'Order �� 30�� ������ ����
    
    SSPan = "DATA ���� ������Դϴ�."
    SSPan.ForeColor = &H0&                  ' black  &H00000000&

Exit Sub

ErrCancel:
    MsgBox "�ڷ������ ����Ͽ����ϴ�."
    Call ErrCancel_Click                    ' �ڷ���� ����
Exit Sub


ErrorMsg:
    MsgBox "Error " & "Code = " & Err.Number & vbLf & vbLf & Err.Description
    If FileClose = False Then
        Close #hSaveFile                    ' error �߻��� file close
    End If


'    SS.Enabled = True
    Timer_Picture.Interval = 0              'Timer_Picture_Timer End
    Timer_Order.Interval = 0              'Timer_order_Timer End
    timerx = False                          '����ǥ�� ������ flag
    Image1(0).Visible = False               '����ǥ�� image
    Image1(1).Visible = False               '����ǥ�� image
    Image1(2).Visible = False               '����ǥ�� image
    Image1(3).Visible = False               '����ǥ�� image
    Image1(4).Visible = False               '����ǥ�� image
    Call WorkDisplay(0)
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If

End Sub


'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}

Private Sub MSComm1_OnComm() ' DATA ���� ó��
    Dim EVMsg$
    Dim ERMsg$
    Select Case MSComm1.CommEvent           'CommEvent �Ӽ��� ���� �׸�
       '�̺�Ʈ �޽���
        Case comEvReceive                   ' ��Ʈ�κ��� �����Ͱ� ������...
             RBuffer = MSComm1.Input
        Case comEvSend
        Case comEvCTS:          EVMsg$ = "CTS ���� ����"
        Case comEvDSR:          EVMsg$ = "DSR ���� ����"
        Case comEvCD:           EVMsg$ = "CD ���� ����"
        Case comEvRing:         EVMsg$ = "��ȭ ���� �︮�� ��"
        Case comEvEOF:          EVMsg$ = "EOF ����"
       '���� �޽���
        Case comBreak:          ERMsg$ = "�ߴ� ��ȣ ����"
        Case comCDTO:           ERMsg$ = "�ݼ��� ���� �ð� �ʰ�"
        Case comCTSTO:          ERMsg$ = "CTS �ð� �ʰ�"
        Case comDCB:            ERMsg$ = "DCB �˻� ����"
        Case comDSRTO:          ERMsg$ = "DSR �ð� �ʰ�"
        Case comFrame:          ERMsg$ = "�����̹� ����"
        Case comOverrun:        ERMsg$ = "�и�Ƽ ����"
        Case comRxOver:         ERMsg$ = "���� ���� �ʰ�"
        Case comRxParity:       ERMsg$ = "�и�Ƽ ����"
        Case comTxFull:         ERMsg$ = "���� ���ۿ� ������ ����"
        Case Else:              ERMsg$ = "�� �� ���� ���� �Ǵ� �̺�Ʈ"
    End Select
    
    ' error message ���
    If ERMsg <> "" And FileClose = False Then
            SSPan = "Error  " & ERMsg$
            ErrList.AddItem "Error  " & ERMsg$
            ErrList.ListIndex = ErrList.ListCount - 1
    End If
    
    ' event message ���
    If EVMsg <> "" And FileClose = False Then
            SSPan = "Detect " & EVMsg$
            ErrList.AddItem "Detect " & EVMsg$
            ErrList.ListIndex = ErrList.ListCount - 1
    End If
    
    Select Case RBuffer                                 ' Message Block ������ ���� ���
'           Case SOH:    SOHBuffer = True                ' SOH [] Check & Data ������ Buffer Clear
           Case STX:    STXBuffer = True                ' STX [] Check��
           Case ETX:    ETXBuffer = True                ' ETX [] Check��
           Case EOT:    EOTBuffer = True                ' EOT [] Check��
           Case ENQ:    ENQBuffer = True                ' ENQ     Check��
           Case ACK:    ACKBuffer = True                ' ACK     Check��
           Case NACK:   NACKBuffer = True               ' NACK    Check��
    End Select
    
    RBufferSum = RBufferSum & RBuffer                   ' comm port ���� �Է��� data ����
    
    
    If ENQBuffer = True Then
        If FileClose = False Then
            Print #hSaveFile, "Rx " & Format(lblTime, "hh:mm:ss") & " ]  " & ENQ
        End If
        Call Ack_Send                                   'DATA �۽� ���� Flag from STA
'        ENQBuffer = False
    End If
    
    
    If STXBuffer = True And RBuffer = vbLf Then         'STX & LF  Check
        If FileClose = False Then
            Print #hSaveFile, RSaveRecord(RBufferSum)       ' write omitting cr lf
        End If
        Call DataReceive                                '������ data record �м�
        RBufferSum = ""
        RBuffer = ""
        STXBuffer = False
        Call Ack_Send
    End If
    
    If EOTBuffer = True And GBTransmit = "Q" Then
'    If EOTBuffer = True Then
        If FileClose = False Then
            Print #hSaveFile, "Rx " & Format(lblTime, "hh:mm:ss") & " ]  " & EOT
        End If
        Call Work_List_Return                           'DATA ���� ���� Flag
        EOTBuffer = False
    End If

    
    If ACKBuffer = True Then
        If FileClose = False Then
            Print #hSaveFile, "Rx " & Format(lblTime, "hh:mm:ss") & " ]  " & ACK
        End If
        Call Order_Data_Send                            'DATA �۽� ���� Flag
        ACKBuffer = False
    
    End If
    
End Sub


Private Sub DataReceive()

    On Error Resume Next

    ' Message Block
    ' rbuffersum   Data Format Sample    @ : <CR>
    ' rbuffersum   Data Format Sample    # : <LF>
    ' rbuffersum   Data Format Sample     : <SOH>
    ' rbuffersum   Data Format Sample     : <STX>
    ' rbuffersum   Data Format Sample     : <ETX>
    ' rbuffersum   Data Format Sample     : <EOT>
    '
    
    Dim Msequence
    Dim Mresult
    Dim Malarm
    
    
    'Header Check & Move
    If Mid(RBufferSum, 3, 1) = "H" Then
        RCheckSum = RBufferSum
        RCheckSumD = Checksum_Eci_Rx(RCheckSum)

        If RCheckSumD <> Mid(RBufferSum, 36, 2) Then
            ErrList.AddItem " Header CheckSum Error "
            ErrList.ListIndex = ErrList.ListCount - 1
        End If
        
        Mdate = Mid(RBufferSum, 20, 14)    ' yyyymmddhhmmss
        
'        If Mid(RBufferSum, 25, 1) = "Q" Then
'            HeaderQC = True
'        ElseIf Mid(RBufferSum, 25, 1) = "P" Then
'            HeaderQC = False
'        End If
        
    End If
    
    
    'Check Patient File or Work List
    If Mid(RBufferSum, 2, 1) = 2 And Mid(RBufferSum, 3, 1) = "P" Then       'Transmition of Patient File to STA
        GBTransmit = "P"
    End If
    
    '---------------------------------------------------------------------------
    '   1) Transmition of Patient File to STA
    '---------------------------------------------------------------------------
    If GBTransmit = "P" Then
        'Request for a Working List from STA
        
        SSPan.ForeColor = &HFF&                         ' red  &H000000FF&
        SSPan = "Patient Result Data �������Դϴ�..........."
        
        Select Case Mid(RBufferSum, 3, 1)
               'Patient information Record from interface
               Case "P"
                      RCheckSum = RBufferSum
                      RCheckSumD = Checksum_Eci_Rx(RCheckSum)

                     If RCheckSumD = Mid(RBufferSum, 22, 2) Then
               
                          For i = 12 To Len(RBufferSum)
                              If Mid(RBufferSum, i, 1) = vbCr Then
                                  Pinfo1 = Mid(RBufferSum, 12, i - 12)
                                  Exit For
                              End If
                          Next i
                         
                     Else
                         ErrList.AddItem " CheckSum Error !!!!!!"
                         ErrList.AddItem " Pointer : P "
                         ErrList.ListIndex = ErrList.ListCount - 1
                         Exit Sub
                     End If
                     
                     RJeobsuNo_Error = False
                     
               'Test order Record Specimen ID
               Case "O"    'oh
                      RCheckSum = RBufferSum
'                      RCheckSumD = CheckSum_STX(RCheckSum)
                      RCheckSumD = Checksum_Eci_Rx(RCheckSum)


                     If RCheckSumD = Mid(RBufferSum, Len(RBufferSum) - 3, 2) Then
                         
                         
                          For i = 7 To Len(RBufferSum)
                              If Mid(RBufferSum, i, 1) = "^" Then
                                  RJeobsuNo = Mid(RBufferSum, 7, i - 7)
                                  Exit For
                              End If
                          Next i
                         
                         If Len(RJeobsuNo) <> 12 Then
                            ErrList.AddItem " RJeobsuNo Length Error !!!!!!"
                            ErrList.AddItem " Pointer : O "
                            ErrList.ListIndex = ErrList.ListCount - 1
                            
                            Pinfo1 = ""
                            RJeobsuNo_Error = True

                            Exit Sub
                            
                         End If
                         
                         Optno = Pinfo1
                         
                         If Mid(RBufferSum, Len(RBufferSum) - 29, 1) = "S" Then
                             StrGBER = "S"
                         ElseIf Mid(RBufferSum, Len(RBufferSum) - 29, 1) = "R" Then
                             StrGBER = "R"
                         Else
                             StrGBER = "R"
                         End If
                         
                         '1)�������� Order ���� ���۽� ���
'                         For i = QCounter To 1 Step -1
                         For i = SS.DataRowCnt To 1 Step -1
                             If Temp_Jeobsu(i, 1) = RJeobsuNo Then
                                 RPoint = i
                                 Temp_Result(i, 3) = Optno
                                Exit For
                             End If
                         Next i
                         
                         '����� data Ȯ��
                         For i = 1 To RPoint - 1
                             If Temp_Jeobsu(i, 1) = RJeobsuNo Then
                                 SS.Row = i
                                 For j = 1 To 6
                                     SS.Col = j
                                     SS.ForeColor = RGB(255, 237, 250)
                                 Next j
                                 Exit For
                             End If
                         Next i
                         
                         
                         '2)batch�� ��� data ���Ž� ���
                         If strBiDirect_Trans = False Then
                             QCounter = QCounter + 1
                             RPoint = QCounter
                             Temp_Jeobsu(RPoint, 1) = RJeobsuNo
                         End If
                         
                        If StrGBER = "S" Then
                            SS.Row = RPoint
                            For i = 1 To 6
                                SS.Col = i
                                SS.ForeColor = RGB(255, 0, 0)
                            Next i
                            StrGBER = "R"
                        End If
                         
'                         Call Ack_Send
                     
                     Else
                         ErrList.AddItem " CheckSum Error !!!!!!"
                         ErrList.AddItem " Pointer : O "
                         ErrList.ListIndex = ErrList.ListCount - 1
                         Exit Sub
                     End If
                     
               'Result Record
               Case "R"
                      If RJeobsuNo_Error = True Then
                          Exit Sub
                      End If

                      
                      RCheckSum = RBufferSum
                      RCheckSumD = Checksum_Eci_Rx(RCheckSum)
                     
                      If RCheckSumD = Mid(RBufferSum, Len(RBufferSum) - 3, 2) Then
                         
                          RSequence = Mid(RBufferSum, 5, 1)
                         
                          'if rsequence Error then NAK Send
                         
                          RResult = Mid(RBufferSum, 19, 3)    '�˻� �ڵ�
                          
                          
                          For i = 1 To 60
                              If Temp_K(i, 3) = RResult Then
                                  CPoint = Val(Temp_K(i, 2))
                                  Exit For
                              End If
                          Next i
                          
                          '����� ����
                          If strBiDirect_Trans = False Then
                              For i = 1 To 60
                                  If Temp_K(i, 3) = RResult Then
                                      Temp_Jeobsu(RPoint, CPoint) = Trim(Temp_K(i, 0))
                                      Exit For
                                  End If
                              Next i
                          End If
                          
                         
                          If Mid(RBufferSum, 25, 2) = "No" Then
                              ErrList.AddItem " No Result : MEPFRE => => => Error "
                              Exit Sub
                          End If
                          
                          For i = 25 To Len(RBufferSum)
                              If Mid(RBufferSum, i, 1) = "|" Then
                                  Temp_Result(RPoint, CPoint) = Mid(RBufferSum, 25, i - 25)
                                  R_Check = True
                                  Exit For
                              End If
                          Next i
                          
                          Dim SPCounterS     As Integer    ' error code ������ġ check��
                          Dim SPCounterE     As Integer    ' error code ������ġ check

                          For i = 1 To Len(RBufferSum)
                              If Mid(RBufferSum, i, 1) = "^" Then
                                  SPCounterS = SPCounterS + 1
                                  If SPCounterS = 5 Then
                                      SPCounterS = i
                                      Exit For
                                  End If
                              End If
                          Next i
                         
                         
                          For i = 1 To Len(RBufferSum)
                              If Mid(RBufferSum, i, 1) = "|" Then
                                  SPCounterE = SPCounterE + 1
                                  If SPCounterE = 7 Then
                                      SPCounterE = i
                                      Exit For
                                  End If
                              End If
                          Next i
                          
                          If (SPCounterE - SPCounterS) > 1 Then
                              Dim DataMessage   As String
                              If Mid(RBufferSum, SPCounterS, SPCounterE - SPCounterS) <> "M1" Then
                                  ErrList.AddItem "Data Message :" & Mid(RBufferSum, SPCounterS, SPCounterE - SPCounterS)
                              End If
                          End If
                          Fcheck = 0
                      Else
                          ErrList.AddItem " CheckSum Error !!!!!!"
                          ErrList.AddItem " Pointer : R "
                          ErrList.ListIndex = ErrList.ListCount - 1
                          Exit Sub
                      End If
               
               Case "L"
                      If RJeobsuNo_Error = True Then
                          Exit Sub
                      End If
                      
                      RCheckSum = RBufferSum
'                      RCheckSumD = CheckSum_STX(RCheckSum)
                      RCheckSumD = Checksum_Eci_Rx(RCheckSum)
                      
                      If RCheckSumD = Mid(RBufferSum, 10, 2) Then
                          Lterminator = Mid(RBufferSum, 7, 1)
                          Select Case Lterminator
                                 Case "N"
                                       LNormal = True
                                      'normal termination
                                 Case "T"
                                       ErrList.AddItem " Sender aborted"
                                       ErrList.ListIndex = ErrList.ListCount - 1
                                 Case "R"
                                       ErrList.AddItem " Receiver requested abort "
                                       ErrList.ListIndex = ErrList.ListCount - 1
                                 Case "E"
                                       ErrList.AddItem " Unknown system error "
                                       ErrList.ListIndex = ErrList.ListCount - 1
                                 Case "Q"
                                       ErrList.AddItem " Error in last request for information "
                                       ErrList.ListIndex = ErrList.ListCount - 1
                                 Case "I"
                                       ErrList.AddItem " No information available from last query "
                                       ErrList.ListIndex = ErrList.ListCount - 1
                                 Case "F"
                                       ErrList.AddItem " Last request for information processed "
                                       ErrList.ListIndex = ErrList.ListCount - 1
                          End Select
                          
                      Else
                          ErrList.AddItem " CheckSum Error !!!!!!"
                          ErrList.AddItem " Pointer : P L "
                      End If
           
        End Select   ' "P" Check and Select
        
        
        'DATA Write to Server
        
        If R_Check = True Then
            If strBiDirect_Trans = True Then
                'Save_Result(PTJeobsuNo, itemcd, Result)
                
                Call Save_Result(Temp_Jeobsu(RPoint, 1), Temp_Jeobsu(RPoint, CPoint), Temp_Result(RPoint, CPoint))
            
            Else
'                Dim AAA
                
'                AAA = S_itemcd(RResult)    'test�� move
'                RPoint = QCounter          'error  check
                
'                Call Save_Result(RJeobsuNo, AAA, Temp_Result(RPoint, CPoint))
                Call Save_Result(Temp_Jeobsu(RPoint, 1), Temp_Jeobsu(RPoint, CPoint), Temp_Result(RPoint, CPoint))
            
            End If
            
            R_Check = False
            ME_Check = False
            MA_Check = False
        
        Else
        
        End If
        '---------------------------------------------------------------------------
        '    End Receive Patient File
        '---------------------------------------------------------------------------
    End If
    
    
    If EOTBuffer = True Then
        RSequence = ""
        RResult = ""
        
        Msequence = ""
        Mresult = ""
        Malarm = ""
        
        STXBuffer = False
        ETXBuffer = False
        EOTBuffer = False
     
        Receive_STA_Check = False                        ' STA Code Check
    
    End If
    

End Sub


'Private Sub Save_Result(ByVal Rresult1 As String, ByVal itemcd1 As String, ByVal Optno1 As String)
Private Sub Save_Result(JeobsuPT, itemcd1, ResultU)
    
    Dim Bdt
    Dim Bno1
    Dim Bno2
        
    Bdt = convLabnoToExpand(Mid(JeobsuPT, 1, 5))
    Bno1 = Mid(JeobsuPT, 6, 2)
    Bno2 = Mid(JeobsuPT, 8, 5)
    
    adoConnect.BeginTrans                          ' TRANSACTION�� ����ÿ� COMMITTRANS�� ������
    
    strSQL = ""
    strSQL = strSQL & "UPDATE TWEXAM_GENERAL_SUB "
    strSQL = strSQL & "   SET RESULT1  =   '" & ResultU & "'"
    strSQL = strSQL & " WHERE JEOBSUDT =  TO_DATE('" & Bdt & "','YYYY-MM-DD') "    '�Էµ� ���ڷ� �˻�
    strSQL = strSQL & "   AND SLIPNO1  =   '" & Bno1 & "' "                        '����
    strSQL = strSQL & "   AND SLIPNO2  =   '" & Bno2 & "' "                        '����
    strSQL = strSQL & "   AND ITEMCD   =    '" & itemcd1 & "'"                     'ITEMCODE
    strSQL = strSQL & "   AND VERIFY   =  'N'"                                    ' ����������� VERIFY OK�Ѱ�쿡�� UPDATE��������
    
    Result = AdoExecute(strSQL)
    If Result = True And Rowindicator > 0 Then
'        SSPan = "DATABASE�� ���� �Ǿ����ϴ�. ( " & RecordCountSum & " ��)"
        SSPan = "DATABASE�� ���� �Ǿ����ϴ�. "
        
        adoConnect.CommitTrans                                                   ' TRANSACTION ����ÿ� COMMIT ��Ŵ
    Else
        ErrList.AddItem "    Verify Data       " & JeobsuPT
        ErrList.AddItem "    or Update Error   " & itemcd1 & "  " & ResultU
        ErrList.ListIndex = ErrList.ListCount - 1
        
        'file write routine insert
        
        adoConnect.RollbackTrans                                                 ' TRANSACTION ERROR�� ROLLBACK ��Ŵ
        SSPan = "DATABASE�� ������ ERROR�� �߻��Ͽ����ϴ�." & vbCrLf & _
                "VERIFY�� DATA���� Ȯ���Ͻʽÿ�."
    End If
    
    
    If strBiDirect_Trans = True Then
        SS.Row = Val(RPoint)
        SS.Col = Val(CPoint) - 4
        SS.Text = Temp_Result(RPoint, CPoint)
    
'        SS.Row = Val(RPoint)                    'test
'        SS.Col = Val(CPoint) - 4                'test
'        SS.Text = Temp_Result(1, CPoint)        'test
        
        SS.Row = Val(RPoint)
        SS.Col = 6
        SS.Text = Val(SS.Text) + 1
        
        Dim Comp_Check
        
        Comp_Check = SS.Text
        
        SS.Col = 5
        
        If Comp_Check = SS.Text Then
            SS.Col = 6
            SS.BackColor = RGB(0, 255, 0)
            
            'general_sub update
            Call Save_Result_Flag(JeobsuPT)
        
        ElseIf Comp_Check > SS.Text Then
            SS.Col = 6
            SS.BackColor = RGB(255, 255, 0)
        End If
    Else
        
        SS.Row = Val(RPoint)
        
        SS.Col = 1
        SS.Text = JeobsuPT
        
        SS.Col = Val(CPoint) - 4
        SS.Text = ResultU
    
        SS.Col = 6
        SS.Text = Val(SS.Text) + 1
        
    End If
    
    
End Sub


Private Sub Work_List_Return()
    Dim SendBuff            As String
    
    SSPan.ForeColor = &HFF&                         ' red  &H000000FF&
    SSPan = "Patient Order Data �۽����Դϴ�..........."
    
    If LNormal = True Then
        SendBuff = ENQ
        If MSComm1.PortOpen = True Then
            MSComm1.Output = SendBuff
            Print #hSaveFile, "Tx " & Format(lblTime, "hh:mm:ss") & " ]  " & SendBuff
        End If
        ENQBuffer = False
    End If

End Sub


Private Sub Ack_Send()
    Dim SendBuff            As String
    
    SendBuff = ACK
    
    If MSComm1.PortOpen = True Then
        MSComm1.Output = SendBuff
        Print #hSaveFile, "Tx " & Format(lblTime, "hh:mm:ss") & " ]  " & SendBuff
    End If
    
    ENQBuffer = False
    RBufferSum = ""
    RBuffer = ""
    
End Sub


Private Sub Order_Data_Send()

'******************** Part 6 **********************************
'*      DB���� �˻��� item code�� STA�� �˻� �䱸 Data ����  *
'**************************************************************
' STA�� �˻�䱸 ITEM CODE��  ����
    'Data type format
    '         1         2         3         4         5         6
    
    Dim SendBuffD           As String           'data
    
    Dim Lencheck
    Dim Specimen            As String

    '--- Buffer Setting ------------------------------------------------

    If SS.DataRowCnt < 1 Then Exit Sub
    
    If RCounter > SS.DataRowCnt Then Exit Sub
    
    
    R1 = RCounter
    SS.Row = R1
    SS.Col = 1
    
    Temp_Jeobsu(R1, 1) = Trim(SS.Text)
    
    SendTime = Format(SysDate_Get, "yyyymmdd") & Format(Time, "hhmmss")
    
        Select Case Or_Seq
               Case 0   ' Send Header
                        SendBuffT = ENQ
                        
                        If MSComm1.PortOpen = True Then
                            MSComm1.Output = SendBuffT               'data send to com port
                            Print #hSaveFile, "Tx " & Format(lblTime, "hh:mm:ss") & " ]  " & SendBuffT
                        End If
               
                        Or_Seq = Or_Seq + 1
               
                        SSPan.ForeColor = &HFF&                         ' red  &H000000FF&
                        SSPan = "ORDER DATA �������Դϴ�..........."
               
               Case 1   ' Send Header
                        SendBuffW = Or_Seq & "H|\^&||||||||||||" & SendTime & vbCr & ETX
                        SendBuffT = STX & SendBuffW & CheckSum_ECi_Tx(SendBuffW) & vbCr & vbLf
                        
                        Call ini_check
                        
                        For C1 = 11 To 50
                            If Temp_Jeobsu(R1, C1) <> "" Then
'                                Temp_Request_Code(R1, C1) = C1
'                                Temp_Request_Code(R1, C1) = Conv_Order_Code(Temp_Jeobsu(R1, C1))
                            
                                For i = 1 To 60
                                    If Temp_K(i, 0) = Temp_Jeobsu(R1, C1) Then
                                        Temp_Request_Code(R1, C1) = Temp_K(i, 3)
                                        Exit For
                                    End If
                                Next i
                            
                            End If
                        Next C1
               
                        Or_Seq = Or_Seq + 1
               
               Case 2   ' Send Patient Information
                        SendBuffW = Or_Seq & "P|1||||||" & Temp_Jeobsu(R1, 3) & vbCr & ETX
                        SendBuffT = STX & SendBuffW & CheckSum_ECi_Tx(SendBuffW) & vbCr & vbLf
                        
                        Or_Seq = Or_Seq + 1
               
               Case 3   ' Send Order Record
                        SendBuffD = ""
                        
                        SendBuffD = "^^^1.000000+"
                        
                        For C1 = 11 To 50
                            If Trim$(Temp_Request_Code(R1, C1)) <> "" Then
                                SendBuffD = SendBuffD & Trim$(Temp_Request_Code(R1, C1)) & "+1" & "\"
                                
                                If Temp_Request_Code(R1, C1) = "044" Then    'NTX
                                    Specimen = "6"
                                Else
                                    Specimen = "4"
                                End If
                            
                            End If
                        Next C1
                        
                        
                        'For i = 1 To 50
                        '    Debug.Print i & "  " & Temp_Request_Code(1, i)
                        '
                        'Next i
                        
                        
                        
                        
                        
                        Lencheck = Len(SendBuffD)
                        
                        If Lencheck = 0 Then
                            ErrList.AddItem " ������ Data�� �����ϴ�. " & Temp_Jeobsu(R1, 1)
                            ErrList.ListIndex = ErrList.ListCount - 1
                            Or_Seq = 0
                            'Exit Function
                        End If
                        
                        SendBuffD = Mid(SendBuffD, 1, Lencheck - 1)
                        
                        ' R : Routine (StrGBER)
                        ' N : New
                        ' 4 : Serum
                        ' 5 : plasma
                        ' 6 : urine             '�˻� : NTX
                        ' 7 : blood
                        ' 8 : amino

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                        'specimen ���� ���� routine �߰�
                        'specimen ���� ���� routine �߰�
                        'specimen ���� ���� routine �߰�
                        'specimen ���� ���� routine �߰�
                        
                        SendBuffW = Or_Seq & "O|1|" & Temp_Jeobsu(R1, 1) & "^^||" & SendBuffD & "|" & StrGBER & "||||||N||||" & Specimen & "||||||||||" & vbCr & ETX
                        
                        SendBuffT = STX & SendBuffW & CheckSum_ECi_Tx(SendBuffW) & vbCr & vbLf
               
                        Or_Seq = Or_Seq + 1
               
               Case 4   ' Send Message Terminator
                        SendBuffW = Or_Seq & "L|1|N" & vbCr & ETX
                        SendBuffT = STX & SendBuffW & CheckSum_ECi_Tx(SendBuffW) & vbCr & vbLf
               
                        Or_Seq = Or_Seq + 1
               Case 5   ' Send EOT
                     
                        SendBuffT = EOT
                        
                        If MSComm1.PortOpen = True Then
                            MSComm1.Output = SendBuffT               'data send to com port
                            Print #hSaveFile, "Tx " & Format(lblTime, "hh:mm:ss") & " ]  " & SendBuffT
                        End If
                     
                        SSPan = "Patient Result ���� ��� ���Դϴ�."
                        
                        Order_Data_Seq = Order_Data_Seq + 1
                        Or_Seq = 0
                     
                        RCounter = RCounter + 1
        
                        SSPan.ForeColor = &H0&                       ' black  &H000000FF&
                        SSPan = "ORDER DATA ���Ŵ�����Դϴ�..........."
        
        End Select
        
        
        If MSComm1.PortOpen = True And SendBuffT <> EOT And SendBuffT <> ENQ Then
            MSComm1.Output = SendBuffT               'data send to com port
'            Print #hSaveFile, "Tx " & Format(lblTime, "hh:mm:ss") & " ]  " & _
'                               Mid(SendBuffT, 1, (Len(SendBuffT) - 2))
            Print #hSaveFile, TSaveRecord(SendBuffT)
        End If
    
End Sub


Private Sub ini_check()
    Dim Rs                  As ADODB.Recordset
    
    Dim Bjeobsudt
    Dim Bslipno1
    Dim Bslipno2
    
    Dim Verify_Check
'******************** Part 1 **********************************
'*      Spread�� pt no�� slipno2�� Temp_Jeobsu Array�� Move   *
'**************************************************************
     
    R1 = RCounter
        C1 = 1
        
        SS.Row = R1
        SS.Col = C1
        SS.Text = Temp_Jeobsu(R1, C1)
        
                       'Rn  Cn
        Bjeobsudt = convLabnoToExpand(Mid(Temp_Jeobsu(R1, C1), 1, 5))
        Bslipno1 = Mid(Temp_Jeobsu(R1, C1), 6, 2)
        Bslipno2 = Mid(Temp_Jeobsu(R1, C1), 8, 5)
            
        For C1 = 2 To 4
            SS.Row = R1
            Select Case C1
                   Case 2   'date
                        SS.Col = C1
                        Temp_Jeobsu(R1, C1) = Bjeobsudt
                        SS.Text = Temp_Jeobsu(R1, C1)
                   
                   Case 3   'PTNO
                        SS.Col = C1
                        Temp_Jeobsu(R1, C1) = PTNOSearch(Temp_Jeobsu(R1, 1))
                        SS.Text = Temp_Jeobsu(R1, C1)
                   Case 4   'Name
                        SS.Col = C1
                        Temp_Jeobsu(R1, C1) = NameSearch(Temp_Jeobsu(R1, 3))
                        SS.Text = Temp_Jeobsu(R1, C1)
            End Select

        Next C1
        
        strSQL = ""
        strSQL = strSQL & " SELECT ITEMCD, GEOMJAN2, GEOMJAN3,GBER "
        strSQL = strSQL & "   FROM TWEXAM_GENERAL_SUB A, "                   ' �˻�������� ���λ���
        strSQL = strSQL & "        TWEXAM_ITEMML B, "                        ' �˻� ITEM MASTER
        strSQL = strSQL & "        TWEXAM_GENERAL C "                        ' �˻��������
        strSQL = strSQL & "  WHERE A.JEOBSUDT = TO_DATE('" & Bjeobsudt & "','YYYY-MM-DD')"
        strSQL = strSQL & "    AND A.SLIPNO1 =   '" & Bslipno1 & "'"        ' �Ϸù�ȣ
        strSQL = strSQL & "    AND A.SLIPNO2 =   '" & Bslipno2 & "'"        ' �Ϸù�ȣ
        strSQL = strSQL & "    AND A.ITEMCD = B.CODEKY "
        strSQL = strSQL & "    AND B.GBROUTINE = 'I'   "
        strSQL = strSQL & "    AND A.PTNO = C.PTNO "
        strSQL = strSQL & "    AND A.JEOBSUDT = C.JEOBSUDT "
        strSQL = strSQL & "    AND A.SLIPNO1 = C.SLIPNO1 "
        strSQL = strSQL & "    AND A.SLIPNO2 = C.SLIPNO2 "
        strSQL = strSQL & "    AND B.GEOMJAN1 = '" & GGJCODE & "' "
        
        Result = AdoOpenSet(Rs, strSQL)
        
        'Debug.Print Rowindicator
        
        If Result Then
            Do While Not Rs.EOF
                If Val(Trim(Rs.Fields("GEOMJAN2") & "")) >= "11" Then
                    Temp_Jeobsu(R1, Val(Trim(Rs.Fields("GEOMJAN2") & ""))) = Trim(Rs.Fields("ITEMCD") & "")
                    If Trim(Rs.Fields("GBER") & "") = "E" Then
                        StrGBER = "S"
                    Else
                        StrGBER = "R"
                    End If
                
                End If

                Rs.MoveNext
            Loop
        End If
        
    If StrGBER = "S" Then
        SS.Row = R1
        For i = 1 To 6
            SS.Col = i
            SS.ForeColor = RGB(255, 0, 0)
        Next i
    End If
    
    SS.Col = 5
    Temp_Jeobsu(R1, 5) = Rowindicator
    SS.Text = Rowindicator

    SS.SetFocus                             ' cell active ���·� ����
    SS.Action = SS_ACTION_ACTIVE_CELL       ' ������ ��ġ�� cursor �̵�

End Sub


Private Sub Save_Result_Flag(JeobsuPT2)

    Dim Bdt
    Dim Bno1
    Dim Bno2
        
    Bdt = convLabnoToExpand(Mid(JeobsuPT2, 1, 5))
    Bno1 = Mid(JeobsuPT2, 6, 2)
    Bno2 = Mid(JeobsuPT2, 8, 5)
    
    strSQL = ""
    strSQL = strSQL & " SELECT JEOBSUDT, SLIPNO1, SLIPNO2, STATUS "
    strSQL = strSQL & "   FROM TWEXAM_GENERAL"
    strSQL = strSQL & "  WHERE JEOBSUDT = TO_DATE('" & Bdt & "','YYYY-MM-DD')"
    strSQL = strSQL & "    AND SLIPNO1 =   '" & Bno1 & "'"
    strSQL = strSQL & "    AND SLIPNO2 =   '" & Bno2 & "'"
    strSQL = strSQL & "    AND (STATUS  = 'R' OR STATUS = 'U') "
    
    Result = AdoOpenSet(Rs, strSQL)
        
    If Result Then
        adoConnect.BeginTrans                          ' TRANSACTION�� ����ÿ� COMMITTRANS�� ������
        
        strSQL = ""
        strSQL = strSQL & "UPDATE TWEXAM_GENERAL "
        strSQL = strSQL & "   SET STATUS   = 'U' "
        strSQL = strSQL & " WHERE JEOBSUDT = TO_DATE('" & Bdt & "','YYYY-MM-DD') "    '�Էµ� ���ڷ� �˻�
        strSQL = strSQL & "   AND SLIPNO1  = '" & Bno1 & "' "                        '����
        strSQL = strSQL & "   AND SLIPNO2  = '" & Bno2 & "' "                        '����
        strSQL = strSQL & "   AND (STATUS  = 'R' OR STATUS = 'U') "
        
        Result = AdoExecute(strSQL)
        If Result = True And Rowindicator > 0 Then
            SSPan = "DATABASE�� ���� �Ǿ����ϴ�. "
            adoConnect.CommitTrans                                                   ' TRANSACTION ����ÿ� COMMIT ��Ŵ
        Else
            ErrList.AddItem "    Verify OK         " & JeobsuPT2
            ErrList.AddItem "    or �������       "
            ErrList.ListIndex = ErrList.ListCount - 1
            
            adoConnect.RollbackTrans                                                 ' TRANSACTION ERROR�� ROLLBACK ��Ŵ
            SSPan = " DATABASE�� ������ ERROR�� �߻��Ͽ����ϴ�." & vbCrLf & _
                    " ����Ϸ�� DATA���� Ȯ���Ͻʽÿ�."
        End If
        
    End If

End Sub

Private Sub mnuEnd_Click()
'2)��������
    
    On Error Resume Next
'    If SS.DataRowCnt < 1 Then Exit Sub

    If Receive_Check = False Then Exit Sub
    
    
    Dim Msg, Style, Title, Response
    RecordCount = 0
    Msg = " �˻縦 ���� �Ͻðڽ��ϱ�?" & vbCrLf & " �̼��ŵ� �ڷḦ Ȯ���ϼ̽��ϱ�?"
    Style = vbYesNo + vbQuestion + vbDefaultButton2     ' Define buttons.
    Title = "�˻縦 ���� Ȯ��"                          ' �⺻ ����.
    Response = MsgBox(Msg, Style, Title)
    
    If Response = vbNo Then Exit Sub                    ' ����ڰ� �ƴϿ� ���ý� ������.
    
    Receive_Check = False
        
    For i = 1 To SS.MaxRows
        For j = 1 To 6
            SS.Row = i
            SS.Col = j
            SS.Lock = False
        Next j
    Next i
    
'        SS.Enabled = True
    Timer_Picture.Interval = 0                       'Timer_Picture_Timer End
    Timer_Order.Interval = 0                       'Timer_order_Timer End
    timerx = False
    Image1(0).Visible = False
    Image1(1).Visible = False
    Image1(2).Visible = False
    Image1(3).Visible = False
    Image1(4).Visible = False
    Receive_STA_Check = False
    
    Call WorkDisplay(0)

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    Print #hSaveFile, " "
    ResultText = ""
    For R1 = 1 To MaxDataRowCnt
        For C1 = 7 To SS.MaxCols + 6
            If Temp_Result(R1, C1) <> "" Then
                ResultText = ResultText & "  " & C1 & " = " & Temp_Result(R1, C1)
            End If
        Next C1
        Print #hSaveFile, Format$(R1, "000") & " " & ResultText
        ResultText = ""
    Next R1
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    Print #hSaveFile, " " & vbCrLf & "@@@@@  Spread Data" & vbCrLf & " "
    ResultText = ""
    For R1 = 0 To MaxDataRowCnt
        For C1 = 1 To SS.MaxCols
            SS.Row = R1
            SS.Col = C1
            If C1 <> 3 Then
                If SS.Text = "" Then SS.Text = "0"
                ResultText = ResultText & Format$(Trim$(SS.Text), "@@@@@@@@@@@@@@") & " : "
            End If
        Next C1
        Print #hSaveFile, ResultText
        ResultText = ""
    Next R1
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    
    If FileClose = False Then
        Close #hSaveFile
    End If
    FileClose = True
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If

End Sub


'�̻��
Private Sub mnuWrite_Click()
'3)�ڷ�����
'Spread sheet�� data�� Server�� ����
    
    If strBiDirect_Trans = True Then
        Exit Sub
    End If
    
    On Error Resume Next
    
    If SS.DataRowCnt < 1 Then Exit Sub
    Dim Msg, Style, Title, Response
    Msg = " �ڷḦ DATABASE�� " & vbCrLf & "�����Ͻðڽ��ϱ�?"
    Style = vbYesNo + vbQuestion + vbDefaultButton2 ' Define buttons.
    Title = "DATABASE UPDATE"                       ' �⺻ ����.
    Response = MsgBox(Msg, Style, Title)
    If Response = vbYes Then                        ' ����ڰ� ���� ����.
        Call Data_Update                            ' �ӻ󺴸� �˻����� 11(��ȭ��)�˻翡���� ITEM CODE �˻� & SET
    End If

End Sub

'�̻�� mnuWrite_Click���� call
Private Sub Data_Update()
    Dim Rs                  As ADODB.Recordset
 
    SSPan = "DATABASE�� �����ϰ� �ֽ��ϴ�."
    Pflag = False
    JeobsuCheck = True

''''''''''''''''''''''''''''''''''''''''''''''' TRANSACTION �� ������ġ ����
    adoConnect.BeginTrans                          ' TRANSACTION�� ����ÿ� COMMITTRANS�� ������
    
    ' DATABASE UPDATE                                                                       '
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '�ӻ󺴸� �˻����� 11(��ȭ�� �ڵ��м�)�˻翡���� ITEM CODE SETTING
    ' Temp_Jeobsu(R1,C1)�� settting �Ǿ�����
    
    For R1 = 1 To MaxDataRowCnt
        For C1 = 7 To MaxRecordCount + 6        'itemcd�� check�ϱ����� for next
            If Trim$(Temp_Result(R1, C1)) <> "" Then
            
'                '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'                List3.AddItem C1 & "  " & Temp_Result(R1, C1)
'                List3.ListIndex = List3.ListCount - 1
'                Debug.Print Temp_Result(R1, C1)
                '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                
                strSQL = ""
                strSQL = strSQL & "UPDATE TWEXAM_GENERAL_SUB1 "
                strSQL = strSQL & "   SET RESULT1  =   '" & Format(Val(Temp_Result(R1, C1)), "###0.000") & "'"
                strSQL = strSQL & " WHERE JEOBSUDT =  TO_DATE('" & Temp_Jeobsu(R1, 4) & "','YYYY-MM-DD') "      '�Էµ� ���ڷ� �˻�
                strSQL = strSQL & "   AND SLIPNO1  =   11"                                                      '����
                strSQL = strSQL & "   AND VERIFY   =  'N'"                                                      ' ����������� VERIFY OK�Ѱ�쿡�� UPDATE��������
                strSQL = strSQL & "   AND SLIPNO2  =     " & Temp_Jeobsu(R1, 2)                                 '���� 2���̻��� ��� CHECK
                strSQL = strSQL & "   AND PTNO     =    '" & Trim$(Temp_Jeobsu(R1, 1)) & "'"                    'PATIENT NUMBER
                strSQL = strSQL & "   AND ITEMCD   =    '" & Trim$(Temp_K(C1 - 6, 0)) & "'"                     'ITEMCODE
                
                Result = AdoExecute(strSQL)
                If Result >= 0 And Rowindicator > 0 Then
                    RecordCountBit = 1
                ElseIf Result = -1 Then
                    MsgBox "Check Error" & vbCrLf & R1 & "��° data�� Ȯ���Ͻʽÿ� "
                    JeobsuCheck = False
                End If
            End If
        Next C1
        RecordCountSum = RecordCountSum + RecordCountBit
        RecordCountBit = 0
    Next R1
          
    If Result Then
        SSPan = "DATABASE�� ���� �Ǿ����ϴ�. ( " & RecordCountSum & " ��)"
        If RecordCountSum = 0 Then SSPan = " ����� Data�� �����ϴ�."
        adoConnect.CommitTrans                                                   ' TRANSACTION ����ÿ� COMMIT ��Ŵ
        RecordCountSum = 0
        Update_Check = False
    Else
        MsgBox "   Update Error     "
        adoConnect.RollbackTrans                                                 ' TRANSACTION ERROR�� ROLLBACK ��Ŵ
        SSPan = "DATABASE�� ������ ERROR�� �߻��Ͽ����ϴ�."
        Update_Check = False
    End If
    
End Sub


Private Sub mnuClear_Click()
'4)ȭ�� Clear
    Dim Msg, Style, Title, Response
    If Update_Check = True Then
        Msg = " ������ DATA�� �������� �ʾҽ��ϴ�." & vbCrLf & _
              " DATA�� �����Ѵ��� ȭ���� CLEAR�Ͻʽÿ�." & vbCrLf & _
              "                                " & vbCrLf & _
              " ȭ���� CLEAR�Ͻðڽ��ϱ�?"
        Style = vbYesNo + vbDefaultButton2 + vbCritical     ' Define buttons.
        Title = "ȭ�� CLEAR"                                ' �⺻ ����.
        Response = MsgBox(Msg, Style, Title)
        
        If Response = vbYes Then                            ' ����ڰ� ���� ����.
            Call SS_INIT(SS, 1, 1)
            Call GotoSpreadSet
            Or_Seq = 0
            SSPan = ""
            Update_Check = False
            ErrList.Clear
            txtBarCode.Text = ""
            txtBarCode.SetFocus
        End If
    Else
        Msg = " ȭ���� CLEAR�Ͻðڽ��ϱ�?"
        Style = vbYesNo + vbQuestion + vbDefaultButton2     ' Define buttons.
        Title = "ȭ�� CLEAR"                                ' �⺻ ����.
        Response = MsgBox(Msg, Style, Title)
        If Response = vbYes Then                            ' ����ڰ� ���� ����.
            Call SS_INIT(SS, 1, 1)
            Call GotoSpreadSet
            Or_Seq = 0
            SSPan = ""
            StrGBER = "R"
            GBTransmit = ""
            txtBarCode.Text = ""
            txtBarCode.SetFocus
            For i = 0 To 100
                For j = 0 To 50
                    Temp_Jeobsu(i, j) = ""
                    Temp_Request_Code(i, j) = ""
                    Temp_Result(i, j) = ""
                Next j
            Next i
        End If
    End If
    
End Sub


Private Sub mnuSet_Click()
'5)���ȯ�漳��
    Dim CodeCheck

    frmSetComm.Show vbModal
    
    CodeCheck = GetSetting("LabInterface", "SetPC", "GGJCODE" & GGJCODE)
    GGCODE = Mid(CodeCheck, 1, 2)
    If Mid(CodeCheck, 6, 1) = "1" Then
        ComPort = GetSetting("LabInterface", "SetComm", "ComPort1")
        Settings = GetSetting("LabInterface", "SetComm", "ComSettings1")
    ElseIf Mid(CodeCheck, 6, 1) = "2" Then
        ComPort = GetSetting("LabInterface", "SetComm", "ComPort2")
        Settings = GetSetting("LabInterface", "SetComm", "ComSettings2")
    End If
    
    If ComPort = "" Then
        MsgBox " COM PORT ������ �߸� �Ǿ����ϴ�."
        Exit Sub
    End If
    
    If MSComm1.PortOpen = False Then
        MSComm1.CommPort = ComPort
        MSComm1.Settings = Settings
    End If
    Call GotoSpreadSet

End Sub


Private Sub mnuExit_Click()
'6)����
    Dim Msg, Style, Title, Response
    If Update_Check = True Then
        Msg = " ������ DATA�� �������� �ʾҽ��ϴ�." & vbCrLf & _
              " DATA�� �����Ѵ��� �����Ͻʽÿ�." & vbCrLf & _
              "                                " & vbCrLf & _
              " ���α׷��� �����Ͻðڽ��ϱ�?"
        Style = vbYesNo + vbDefaultButton2 + vbCritical     ' Define buttons.
        Title = "���α׷� ����"                             ' �⺻ ����.
        Response = MsgBox(Msg, Style, Title)
        If Response = vbYes Then                            ' ����ڰ� ���� ����.
            If MSComm1.PortOpen = True Then
                MSComm1.PortOpen = False
            End If
            Update_Check_Force = True                       '��������� ���
            Unload Me
        End If
    Else
        Msg = " ���α׷��� �����Ͻðڽ��ϱ�?"
        Style = vbYesNo + vbDefaultButton2 + vbCritical     ' Define buttons.
        Title = "���α׷� ����"                             ' �⺻ ����.
        Response = MsgBox(Msg, Style, Title)
        If Response = vbYes Then                            ' ����ڰ� ���� ����.
            If MSComm1.PortOpen = True Then
                MSComm1.PortOpen = False
            End If
            Update_Check_Force = True                       '��������� ���
            Unload Me
        End If
    End If

End Sub


Private Sub lblDate_Click()
    Call FrmCalendar.Calendar_Show(lblDate)
    Call GotoSpreadSet
    
End Sub


Private Sub vaSpread_Display(ResultText, ssR1, ssC1)
    SS.Row = ssR1
    SS.Col = ssC1
    SS.Text = ResultText
    
    SS.Col = 6
    SS.Action = SS_ACTION_ACTIVE_CELL
    
End Sub

Sub vaSpread_Clear(SS, SColumn, SRow, EColumn, ERow)

    SS.Col = SColumn: SS.Col2 = SS.MaxCols
    SS.Row = SRow: SS.Row2 = -1
     
    SS.BlockMode = True
    SS.Action = SS_ACTION_CLEAR_TEXT
    SS.BlockMode = False
    
    SS.Col = 1:     SS.Row = 1
    SS.Action = SS_ACTION_ACTIVE_CELL
    
End Sub



Sub GetIniFile()
    Dim CodeCheck
'Registry ������ġ
'HKEY_CURRENT_USER\Software\VB and VBA Program Settings\LabInterface

'Rack number / Position number Set
'    If (GetSetting("LabInterface", "SetPc", "MaxRCntNo")) = "" Then
'        Call SaveSetting("LabInterface", "SetPc", "MaxRCntNo", "1")
'    End If
'    GnRCntNo = Val(GetSetting("LabInterface", "SetPc", "MaxRCntNo"))              ' register���� serial number get
'
'    If GetSetting("LabInterface", "Setpc", "MaxPCntNo") = "" Then
'        Call SaveSetting("LabInterface", "SetPc", "MaxPCntNo", "1")
'    End If
'    GnPCntNo = Val(GetSetting("LabInterface", "SetPc", "MaxPCntNo"))              ' register���� serial number get
    
    
'���ȯ�� �ʱ⼳�� Ȯ�ι� �⺻ ȯ�� ����
    
    GGJCODE = "13"
    
'    Call SaveSetting("LabInterface", "SetPC", "GGJCODE" & GGJCODE, GGJCODE)
    
    CodeCheck = GetSetting("LabInterface", "SetPC", "GGJCODE" & GGJCODE)
    
    GGCODE = Mid(CodeCheck, 1, 2)
    
    If GGCODE = "" Then
        Call mnuSet_Click
    End If
    
    If Mid(CodeCheck, 6, 1) = "1" Then
        ComPort = GetSetting("LabInterface", "SetComm", "ComPort1")
        Settings = GetSetting("LabInterface", "SetComm", "ComSettings1")
    ElseIf Mid(CodeCheck, 6, 1) = "2" Then
        ComPort = GetSetting("LabInterface", "SetComm", "ComPort2")
        Settings = GetSetting("LabInterface", "SetComm", "ComSettings2")
    End If
    
    If MSComm1.PortOpen = False Then
        MSComm1.CommPort = ComPort
        MSComm1.Settings = Settings
    End If
    
    lblPort.Caption = "Com" & ComPort & "," & Settings
    
End Sub


Sub Parmini()                              '���α׷� �ʱ⼳�� �Ķ����
    Timer1.Interval = 500                  'Timer1_Timer 1sec
    
    lblDate = SysDate_Get
    
    lblDate.Alignment = 2
    lblDate.FontSize = 14
    lblDate.BorderStyle = 1
    
    lblTime.Alignment = 2
    lblTime.FontSize = 14
    lblTime.BorderStyle = 1
    
    SSPan.FontSize = 13     '14
    
    Update_Check = False
    Update_Check_Force = False
    
    End_check = False
    
    '��ſ� Definition Character
    SOH = Chr(1)                '<SOH> []
    STX = Chr(2)                '<STX> []
    ETX = Chr(3)                '<ETX> []
    EOT = Chr(4)                '<EOT> []
    ENQ = Chr(5)                '<ENQ>
    ACK = Chr(6)                '<ACK>
    NACK = Chr(21)              '<NACK>
    ETB = Chr(23)               '<ETB>
    
    Or_Seq = 0
    
End Sub


Sub CodeKy_Search()
    Dim Rs                  As ADODB.Recordset
    
    'code key data �˻�
    strSQL = ""
    strSQL = strSQL & " SELECT CODEKY,YAGEO,GEOMJAN2,GEOMJAN3 "
    strSQL = strSQL & "   FROM TWEXAM_ITEMML "
    strSQL = strSQL & "  WHERE GEOMJAN1 = " & GGCODE                                  'STAii ��� code
    strSQL = strSQL & "  ORDER BY CODEKY ASC "
    
    If AdoOpenSet(Rs, strSQL) Then
        R1 = 0
        Do Until Rs.EOF
            Temp_K(R1 + 1, 0) = Trim$(Rs.Fields("CODEKY") & "")         'CODEKY
            Temp_K(R1 + 1, 1) = Trim$(Rs.Fields("YAGEO") & "")          'YAGEO
            
            Temp_K(R1 + 1, 2) = Trim$(Rs.Fields("GEOMJAN2") & "")       'GEOMJAN2 (Serial number or Position number)
            Temp_K(R1 + 1, 3) = Trim$(Rs.Fields("GEOMJAN3") & "")       'GEOMJAN3 (test code)
            Rs.MoveNext: R1 = R1 + 1
        Loop
        MaxRecordCount = Rowindicator                                   ' �˻��׸� ����
    Else
        MsgBox "CODEKY �˻� ERROR" & vbCrLf & "CODEKY�� �����ϴ�.", vbCritical
    End If
    SS.MaxCols = Rowindicator + 6                                     'Record Count�� check �Ͽ� max columns�� �����Ѵ�.
    
End Sub



Sub WorkDisplay(i)
    Select Case i
        Case 0
           Timer1.Tag = ""
           SSPan = "������ ����Ǿ����ϴ�."
           
           mnuReceive.Enabled = True
           mnuEnd.Enabled = True
           mnuWrite.Enabled = True
           mnuClear.Enabled = True
           mnuSet.Enabled = True
           mnuExit.Enabled = True
           
           Toolbar1.Buttons(1).Enabled = True
           Toolbar1.Buttons(3).Enabled = True
           Toolbar1.Buttons(4).Enabled = True
           Toolbar1.Buttons(5).Enabled = True
           Toolbar1.Buttons(6).Enabled = True
           
'           Frame1.Enabled = True
           
           lblDate.Enabled = True
        Case 1
           Timer1.Tag = "ON"
           SSPan = "DATA ���� ���Դϴ�."
           
           mnuReceive.Enabled = False
'          mnuEnd.Enabled = False
           mnuWrite.Enabled = False
           mnuClear.Enabled = False
           mnuSet.Enabled = False
           mnuExit.Enabled = False
           
           Toolbar1.Buttons(1).Enabled = False
           Toolbar1.Buttons(3).Enabled = False
           Toolbar1.Buttons(4).Enabled = False
           Toolbar1.Buttons(5).Enabled = False
           Toolbar1.Buttons(6).Enabled = False
           
           lblDate.Enabled = False
    End Select
End Sub


Sub Kdelete()

    Dim i
    Dim Wfile               As String
    Dim Kfile               As String
    Dim Kdate               As Date
    Dim Rdate               As Date

    On Error Resume Next
        
        File1.Pattern = "*.*"
        File1.Path = "c:\intdown"
        Kdate = Format(Date, "yyyy-mm-dd") 'Date
        Rdate = Format(Date, "yyyy-mm-dd") 'Date
        
        For i = 1 To 999
            If File1.ListCount = 0 Then
                Exit For
            End If
            File1.ListIndex = File1.ListIndex + 1
            
            Text1.Text = File1.Path & "\" & File1.FileName
            Kfile = Text1.Text
            Wfile = LTrim$(Right$(ExtractTime(Text1.Text), 21))
            If Mid(Wfile, 1, 2) = "20" Then
                Date = Mid(Wfile, 1, 10)
                Kdate = Format(Date, "yyyy-mm-dd") 'Date
            Else
                Date = Mid(Wfile, 1, 8)
                Kdate = Format(Date, "yyyy-mm-dd") 'Date
            End If
            
            If DateAdd("m", 1, Kdate) < Rdate Then                             'file �ۼ� ���� check
                Kill (Kfile)                                       '30�� ������ �ۼ��� file ����
            End If
            
            If File1.ListIndex = File1.ListCount - 1 Then
                File1.ListIndex = 0
                Exit For
            End If
        Next i
        
    Date = Rdate

End Sub


Private Sub Timer1_Timer()
    lblTime = Time

End Sub


Private Sub Timer_Picture_Timer()
    If timerx = False Then Exit Sub
    Tcounter = Tcounter + 1
    Select Case Tcounter
        Case 1
                Image1(0).Visible = True
                Image1(1).Visible = False
                Image1(2).Visible = False
                Image1(3).Visible = False
                Image1(4).Visible = False
                Image1(5).Visible = False
        Case 2
                Image1(0).Visible = True
                Image1(1).Visible = True
                Image1(2).Visible = False
                Image1(3).Visible = False
                Image1(4).Visible = False
                Image1(5).Visible = False
        Case 3
                Image1(0).Visible = True
                Image1(1).Visible = True
                Image1(2).Visible = True
                Image1(3).Visible = False
                Image1(4).Visible = False
                Image1(5).Visible = False
        Case 4
                Image1(0).Visible = True
                Image1(1).Visible = True
                Image1(2).Visible = True
                Image1(3).Visible = True
                Image1(4).Visible = False
                Image1(5).Visible = False
        Case 5
                Image1(0).Visible = True
                Image1(1).Visible = True
                Image1(2).Visible = True
                Image1(3).Visible = True
                Image1(4).Visible = True
                Image1(5).Visible = False
        Case 6
                Image1(0).Visible = True
                Image1(1).Visible = True
                Image1(2).Visible = True
                Image1(3).Visible = True
                Image1(4).Visible = True
                Image1(5).Visible = True
        Case 7
                Image1(0).Visible = False
                Image1(1).Visible = False
                Image1(2).Visible = False
                Image1(3).Visible = False
                Image1(4).Visible = False
                Image1(5).Visible = False
                Tcounter = 0
    End Select

End Sub


Private Sub Timer_Order_Timer()
    If Or_Seq <> 0 Then Exit Sub
    If strBiDirect_Trans = False Then Exit Sub

    Call Order_Data_Send

End Sub

'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
Private Sub Timer_RRequest_Timer()
    
    Chr (124)  '|
    Chr (92)   '\
    Chr (94)   '^
    Chr (38)   '&
    
    Dim SendBuff(8)         As String
    
    '/*     RESULTS REQUEST TRANSFFERD TO STA PATIENT
    
    '--- HEADER BLOCK --------------------------------------------------------
    SendBuff(1) = Chr(1) & vbLf                                     '<SOH><LF>
    SendBuff(2) = "06" & " " & "HOST SYSTEM     " & " " & "09" & vbLf
    '--- DATA BLOCK -------------------------------------------------------
    SendBuff(3) = Chr(2) & vbLf                                     '<STX><LF>
    SendBuff(4) = "10 " & "07" & vbLf
    SendBuff(5) = "11 " & Format(TimerRNo, "000") & "/" & Format(TimerPNo, "00") & vbLf     ' X: RACKNO  Y: POSITION
    SendBuff(6) = "12 " & Format(TimerTCode, "00") & vbLf                               ' Z: �˻�ITEM
    SendBuff(7) = Chr(3) & vbLf                                     '<ETX><LF>
    SendBuff(8) = Chr(4) & vbLf                                     '<EOT><LF>
    
    For i = 1 To 8
        If MSComm1.PortOpen = True Then
            MSComm1.Output = SendBuff(i)
        End If
    Next i
        
    If FileClose = False Then
         For i = 1 To 8
         Print #hSaveFile, "Tx " & Format(lblTime, "hh:mm:ss") & " ]  " & _
                           Mid$(SendBuff(i), 1, (Len(SendBuff(i)) - 1))
         Next i
    End If
        
    'MaxRackNo
    'MaxPosiNo
    TimerTCode = TimerTCode + 1
    If TimerTCode = 8 Then                      ' 02 ~ 07�� 6 ���� �˻� item ' �˻�item�� ������ ��� ������!
        TimerTCode = 2
        TimerPNo = TimerPNo + 1
    End If
    If TimerPNo = 16 Then
        TimerPNo = 1
        TimerRNo = TimerRNo + 1
    End If
    If TimerRNo = MaxRackNo And TimerPNo = Val(MaxPosiNo) + 1 Then
        TimerPNo = 1
        TimerRNo = 1
        Timer_RRequest.Interval = 0
    End If
        
End Sub


Private Sub Dir_Open()
    
    On Error GoTo ErrorMsg
    Ser = Ser + 1
    CommonDialog1.InitDir = "C:\Vitros"
    CommonDialog1.FileName = "V" & Format$(lblDate, "yyyymmdd") & Ser
    
    CommonDialog1.Filter = "All Files (*.*)|*.*|" & _
                           "Text Files" & "(*.txt)|*.txt|"
    CommonDialog1.FilterIndex = 2
    CommonDialog1.CancelError = True                        '���Ȯ���� ���� ���
    
    On Error GoTo ErrCancel
'    CommonDialog1.ShowSave                                  'dialog show
    CommonDialog1.CancelError = True                        'cancel error reset
    On Error GoTo ErrorMsg
    
    DoEvents
    Me.Show
    
    temp_file_Order = CommonDialog1.FileName
    
    Exit Sub

ErrCancel:
    MsgBox "�ڷ������ ����Ͽ����ϴ�."
'    Call ErrCancel_Click                    ' �ڷ���� ����
Exit Sub


ErrorMsg:
    MsgBox "Error " & "Code = " & Err.Number & vbLf & vbLf & Err.Description
    If FileClose = False Then
        Close #hSaveFile                    ' error �߻��� file close
    End If
Return

End Sub


Private Sub File_Open()
    
    On Error GoTo ErrorMsg
    
    hSaveFile_Order = FreeFile
    Open temp_file_Order For Append As hSaveFile_Order
        If Err Then
            MsgBox Error$, vbExclamation
            Close hSaveFile_Order
            hSaveFile_Order = 0
            Exit Sub
        End If
    
Exit Sub
     

ErrorMsg:
    MsgBox "Error " & "Code = " & Err.Number & vbLf & vbLf & Err.Description
    If FileClose = False Then
        Close #hSaveFile_Order
    End If
Return

End Sub


Private Sub txtBarCode_GotFocus()
    txtBarCode.SelStart = 0
    txtBarCode.SelLength = Len(txtBarCode.Text)

End Sub


Private Sub txtBarCode_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim Rs                  As ADODB.Recordset
    
    Dim strPt               As String
    Dim Bjeobsudt           As String
    Dim Bslipno1            As String
    Dim Bslipno2            As String
    
    Dim Checkdouble         As String
    Dim Ptnolen
    Dim Temp_ptno
    Dim Temp_name
    
    Dim Temp_Jeobsu
    
    
    If KeyCode = 13 Then
        If Len(txtBarCode.Text) <> 12 Then
            txtBarCode.SelStart = 0
            txtBarCode.SelLength = Len(txtBarCode.Text)
            txtBarCode.SetFocus
            MsgBox " ����Number Length Error "
            Exit Sub
        End If
        If Len(txtBarCode.Text) = 12 Then
        
            SS.Col = 1
            SS.Row = SS.DataRowCnt
            Temp_Jeobsu = SS.Text
            
            If Temp_Jeobsu = txtBarCode.Text Then
                txtBarCode.SelStart = 0
                txtBarCode.SelLength = Len(txtBarCode.Text)
                txtBarCode.SetFocus
                Exit Sub
            End If
            
            SS.Col = 1
            SS.Row = SS.DataRowCnt + 1
            
            SS.Text = txtBarCode.Text
            strPt = txtBarCode.Text
        
            Bjeobsudt = convLabnoToExpand(Mid(strPt, 1, 5))
            Bslipno1 = Mid(strPt, 6, 2)
            Bslipno2 = Mid(strPt, 8, 5)
            
            SS.Col = 2
            SS.Text = Bjeobsudt

            strSQL = ""
            strSQL = strSQL & " SELECT PTNO "
            strSQL = strSQL & "   FROM TWEXAM_GENERAL_SUB "                  ' �� MASTER
            strSQL = strSQL & "  WHERE JEOBSUDT = TO_DATE('" & Bjeobsudt & "','YYYY-MM-DD')"
            strSQL = strSQL & "    AND SLIPNO1 =   '" & Bslipno1 & "'"        ' �Ϸù�ȣ
            strSQL = strSQL & "    AND SLIPNO2 =   '" & Bslipno2 & "'"        ' �Ϸù�ȣ
            
            Result = AdoOpenSet(Rs, strSQL)
            
            If Result Then
                Rs.MoveFirst
                Do While Not Rs.EOF
                    Temp_ptno = Trim$(Rs.Fields("ptno")) & ""
                    Rs.MoveNext
                Loop
            Else
            
                SS.Text = ""
                SS.SetFocus                             'clear�� cell active ���·� ����
                
                MsgBox " DATABASE�� ��ϵ� ������ ���ų� ������ �߸��Ǿ����ϴ�." & vbCrLf & vbCrLf & _
                       " DATA�� ���Է� �Ͻʽÿ�." & vbCrLf & vbCrLf & _
                       " ���Է� �Ŀ��� ���� ERROR�� �߻��� ��� ����Ƿ� ���� �ٶ��ϴ�."
                Exit Sub
            End If
            
            AdoCloseSet Rs
            
            
            strSQL = ""
            strSQL = strSQL & " SELECT SNAME "
            strSQL = strSQL & "   FROM TWBAS_PATIENT "                  ' �� MASTER
            strSQL = strSQL & "  WHERE PTNO = '" & Temp_ptno & "' "     ' PATIENT NO
            
            If AdoOpenSet(Rs, strSQL) Then
                Rs.MoveFirst
                Do While Not Rs.EOF
                    Temp_name = Trim$(Rs.Fields("sname")) & ""
                    Rs.MoveNext
                Loop
            Else
                SS.Text = ""
                SS.SetFocus                             'clear�� cell active ���·� ����
                SS.Action = SS_ACTION_ACTIVE_CELL
                MsgBox " DATABASE�� ��ϵ� �̸��� ���ų� ������ �߸��Ǿ����ϴ�. " & vbCrLf & _
                       " PTNO�� ���Է� �Ͻʽÿ�. " & vbCrLf & vbCrLf & _
                       " ���Է� �Ŀ��� ��� ���� ERROR�� �߻��� ��� ����Ƿ� ���� �ٶ��ϴ�. "
                Exit Sub
            End If
            
            SS.Col = 3
            SS.Text = Temp_ptno
        
            SS.Col = 4
            SS.Text = Temp_name
            
            SS.Col = 1
            SS.SetFocus                             ' cell active ���·� ����
            SS.Action = SS_ACTION_ACTIVE_CELL       ' ������ ��ġ�� cursor �̵�
            
            txtBarCode.SelStart = 0
            txtBarCode.SelLength = Len(txtBarCode.Text)
            txtBarCode.SetFocus
            
        End If
    Else
    
    End If
    
'    txtBarCode.SelStart = 0
'    txtBarCode.SelLength = Len(txtBarCode.Text)

End Sub



