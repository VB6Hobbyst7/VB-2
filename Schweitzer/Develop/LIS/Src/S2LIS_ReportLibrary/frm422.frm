VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frm422RiPrint 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11160
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frm422.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows �⺻��
   Begin MedControls1.LisLabel LisLabel5 
      Height          =   270
      Left            =   75
      TabIndex        =   3
      Top             =   45
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   476
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "�����н� ����� ��� ����"
      LeftGab         =   100
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   720
      Left            =   75
      TabIndex        =   5
      Top             =   255
      Width           =   10740
      Begin VB.OptionButton optBussDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "�ܷ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B679D&
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   36
         Top             =   390
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton optBussDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B679D&
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   35
         Top             =   150
         Width           =   885
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00DBE6E6&
         Height          =   435
         Left            =   5880
         ScaleHeight     =   375
         ScaleWidth      =   4425
         TabIndex        =   6
         Top             =   180
         Width           =   4485
         Begin VB.OptionButton optPrint 
            BackColor       =   &H00FEF5F3&
            Caption         =   "�������"
            Height          =   375
            Index           =   0
            Left            =   0
            Style           =   1  '�׷���
            TabIndex        =   9
            Top             =   0
            Value           =   -1  'True
            Width           =   1485
         End
         Begin VB.OptionButton optPrint 
            BackColor       =   &H00FFF4FD&
            Caption         =   "�ϰ� �����"
            Height          =   375
            Index           =   1
            Left            =   1485
            Style           =   1  '�׷���
            TabIndex        =   8
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton optPrint 
            BackColor       =   &H00F7F7F7&
            Caption         =   "���� �����"
            Height          =   375
            Index           =   2
            Left            =   2955
            Style           =   1  '�׷���
            TabIndex        =   7
            Top             =   0
            Width           =   1455
         End
      End
      Begin MSComCtl2.DTPicker dtpVfyDt 
         Height          =   375
         Left            =   2475
         TabIndex        =   10
         Top             =   195
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyy-MM-dd"
         Format          =   20840451
         CurrentDate     =   36328
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   2
         Left            =   1275
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   195
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "�� �� �� ��"
         Appearance      =   0
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00DBE6E6&
      Height          =   5925
      Left            =   75
      ScaleHeight     =   5865
      ScaleWidth      =   10680
      TabIndex        =   28
      Top             =   2430
      Width           =   10740
      Begin FPSpread.vaSpread tblOrder 
         Height          =   5835
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   10605
         _Version        =   196608
         _ExtentX        =   18706
         _ExtentY        =   10292
         _StockProps     =   64
         BackColorStyle  =   3
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14411494
         MaxCols         =   20
         MaxRows         =   50
         OperationMode   =   2
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   15463405
         ShadowDark      =   14737632
         SpreadDesigner  =   "frm422.frx":000C
         Appearance      =   1
      End
      Begin FPSpread.vaSpread tblOrdSheet 
         Height          =   5850
         Left            =   0
         TabIndex        =   30
         Top             =   -15
         Width           =   10605
         _Version        =   196608
         _ExtentX        =   18706
         _ExtentY        =   10319
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
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
         GrayAreaBackColor=   14411494
         GridColor       =   14013909
         GridShowVert    =   0   'False
         MaxCols         =   46
         MaxRows         =   20
         OperationMode   =   1
         ScrollBars      =   2
         ShadowColor     =   16252927
         ShadowDark      =   14737632
         ShadowText      =   0
         SpreadDesigner  =   "frm422.frx":0BE0
         TextTip         =   4
      End
      Begin FPSpread.vaSpread tblList 
         Height          =   5550
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Visible         =   0   'False
         Width           =   10605
         _Version        =   196608
         _ExtentX        =   18706
         _ExtentY        =   9790
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         ColHeaderDisplay=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   9
         MaxRows         =   50
         OperationMode   =   1
         ShadowColor     =   15857140
         SpreadDesigner  =   "frm422.frx":1CF9
         UserResize      =   0
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1215
      Left            =   75
      TabIndex        =   11
      Top             =   900
      Width           =   10740
      Begin VB.PictureBox picESign 
         Height          =   500
         Left            =   5805
         ScaleHeight     =   435
         ScaleWidth      =   1140
         TabIndex        =   37
         Top             =   600
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00FEF5F3&
         Caption         =   "��ȸ(&Q)"
         Height          =   510
         Left            =   7425
         Style           =   1  '�׷���
         TabIndex        =   13
         Top             =   435
         Width           =   1320
      End
      Begin VB.CommandButton cmdPreview 
         BackColor       =   &H00FEF5F3&
         Caption         =   "�̸�����(&V)"
         Height          =   510
         Left            =   8745
         Style           =   1  '�׷���
         TabIndex        =   12
         Top             =   435
         Width           =   1320
      End
      Begin VB.Frame fraSetWard 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '����
         Height          =   915
         Left            =   60
         TabIndex        =   14
         Top             =   195
         Width           =   6060
         Begin VB.CheckBox chkAllWard 
            BackColor       =   &H00DBE6E6&
            Caption         =   "��ü����/�����"
            ForeColor       =   &H00C76456&
            Height          =   300
            Left            =   2760
            TabIndex        =   20
            Top             =   135
            Width           =   1725
         End
         Begin VB.CommandButton cmdWardList 
            BackColor       =   &H00DEDBDD&
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2400
            MousePointer    =   14  'ȭ��ǥ�� ����ǥ
            Style           =   1  '�׷���
            TabIndex        =   19
            Top             =   75
            Width           =   315
         End
         Begin VB.TextBox txtWardId 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BackColor       =   &H00F1F5F4&
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1230
            TabIndex        =   18
            Top             =   90
            Width           =   1155
         End
         Begin VB.CommandButton cmdDoctList 
            BackColor       =   &H00DEDBDD&
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2385
            MousePointer    =   14  'ȭ��ǥ�� ����ǥ
            Style           =   1  '�׷���
            TabIndex        =   17
            Top             =   495
            Width           =   330
         End
         Begin VB.TextBox txtDoctId 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BackColor       =   &H00F1F5F4&
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1230
            TabIndex        =   16
            Top             =   495
            Width           =   1140
         End
         Begin VB.CheckBox chkAllDoct 
            BackColor       =   &H00DBE6E6&
            Caption         =   "��ü"
            ForeColor       =   &H00C76456&
            Height          =   300
            Left            =   2775
            TabIndex        =   15
            Top             =   540
            Width           =   705
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   360
            Index           =   0
            Left            =   45
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   90
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   635
            BackColor       =   10392451
            ForeColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Caption         =   "����/�����"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   360
            Index           =   6
            Left            =   45
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   480
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   635
            BackColor       =   10392451
            ForeColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Caption         =   "��ġ��"
            Appearance      =   0
         End
         Begin VB.Label lblDoctNm 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "DoctNm"
            ForeColor       =   &H00734A60&
            Height          =   180
            Left            =   3525
            TabIndex        =   22
            Top             =   600
            Width           =   675
         End
         Begin VB.Label lblWardNm 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "WardNm"
            ForeColor       =   &H00734A60&
            Height          =   180
            Left            =   4500
            TabIndex        =   21
            Top             =   210
            Width           =   720
         End
      End
      Begin VB.Frame fraLabNo 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '����
         Height          =   915
         Left            =   60
         TabIndex        =   23
         Top             =   195
         Visible         =   0   'False
         Width           =   6060
         Begin VB.TextBox txtPtId 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   1260
            TabIndex        =   24
            Text            =   "S00"
            Top             =   225
            Width           =   1275
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   360
            Index           =   1
            Left            =   75
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   225
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   635
            BackColor       =   10392451
            ForeColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Caption         =   "ȯ�� ID"
            Appearance      =   0
         End
         Begin VB.Label lblPtNm 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "ȯ�ڸ�1"
            ForeColor       =   &H00734A60&
            Height          =   180
            Left            =   2610
            TabIndex        =   27
            Top             =   330
            Width           =   630
         End
         Begin VB.Label lblSexAge 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "��/30"
            ForeColor       =   &H00734A60&
            Height          =   180
            Left            =   3525
            TabIndex        =   26
            Top             =   345
            Width           =   450
         End
         Begin VB.Label lblWard 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "61W-111"
            ForeColor       =   &H00734A60&
            Height          =   180
            Left            =   4695
            TabIndex        =   25
            Top             =   360
            Width           =   690
         End
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00EBF3ED&
      Caption         =   "�� ��(&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '�׷���
      TabIndex        =   2
      Tag             =   "0"
      Top             =   8505
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00EBF3ED&
      Caption         =   "��   �� (&P)"
      Height          =   510
      Left            =   6855
      Style           =   1  '�׷���
      TabIndex        =   1
      Tag             =   "0"
      Top             =   8505
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00EBF3ED&
      Caption         =   "ȭ������(&C)"
      Height          =   510
      Left            =   8175
      Style           =   1  '�׷���
      TabIndex        =   0
      Tag             =   "0"
      Top             =   8505
      Width           =   1320
   End
   Begin MedControls1.LisLabel lblPrgBar 
      Height          =   270
      Left            =   75
      TabIndex        =   4
      Top             =   2145
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   476
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "����� ��� ���� ����Ʈ"
      LeftGab         =   100
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '����
      Caption         =   " ���� ��¿��� �Ǽ� :"
      ForeColor       =   &H00404000&
      Height          =   195
      Left            =   240
      TabIndex        =   34
      Top             =   8715
      Width           =   2175
   End
   Begin VB.Label lblCnt 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2340
      TabIndex        =   33
      Top             =   8670
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '����
      Caption         =   " �� ��´���� ����Ʈ���� �����Ͻø� ��� �� ���ܵ˴ϴ�."
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   240
      TabIndex        =   32
      Top             =   8475
      Width           =   5955
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   570
      Index           =   0
      Left            =   90
      Shape           =   4  '�ձ� �簢��
      Top             =   8415
      Width           =   6255
   End
End
Attribute VB_Name = "frm422RiPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event FormClose()

Private objSql       As New clsLISSqlReport
Private strStartDate As String
Private MsgFg As Boolean
Private PtFg As Boolean
Private ClearFg As Boolean

Dim blnLoadChk  As Boolean

Private Sub cmdPrint_Click()
    Dim strRstEntryType As String
    Dim strPtId         As String
    Dim strTestDiv      As String
    Dim strTable        As String
    Dim strSQL          As String
    Dim strImgPath      As String
    Dim i               As Long
    Dim j               As Long
    Dim objProgress     As jProgressBar.clsProgress
    Dim objProgress1    As jProgressBar.clsProgress
    Dim objReport       As clsBatchReport
    Dim strLastDt       As String
    Dim strLastTm       As String
    Dim strPrtDt        As String
    Dim strPrtTm        As String
    Dim lngErrCount     As Long
    
    Dim lngFileNo As Long
    
    lngFileNo = FreeFile
    
    If Printers.Count = 0 Then
        MsgBox "���� ������ �����Ͱ� �����Ƿ� ����� �� �����ϴ�.", vbInformation, "������"
        Exit Sub
    End If
    
    If Not optPrint(2).Value And Trim(txtWardId.Text) = "" Then
        MsgBox "������� ����� ������ �����Ͻʽÿ�.", vbInformation, "��������"
        txtWardId.SetFocus
        Exit Sub
    End If
    
    If Not optPrint(2).Value And Trim(txtDoctId.Text) = "" Then
        MsgBox "��ġ�Ǹ� �����Ͻʽÿ�.", vbInformation, "��ġ�Ǽ���"
        txtDoctId.SetFocus
        Exit Sub
    End If
    
    If lblCnt.Caption = 0 Then
        MsgBox "����� ��� ����Ʈ�� �����ϴ�.", vbInformation, "��� ���"
        Exit Sub
    End If
    
    lngErrCount = 0
    
    MouseRunning
    
    Set objProgress = New jProgressBar.clsProgress
    
    With objProgress
        .Container = Me
        .Left = lblPrgBar.Left + 3
        .Top = lblPrgBar.Top + 3
        .Width = lblPrgBar.Width - 10
        .Height = lblPrgBar.Height - 10
        
'        .SetMyForm Me
'        .Choice = True
'        .Max = tblOrder.MaxRows
'        .Min = 0
'        .Value = 0
'        .XPos = lblPrgBar.Left + 3
'        .YPos = lblPrgBar.Top + 3
'        .XWidth = lblPrgBar.Width - 10 'fraWSHeader.Width - (optCondition(1).Width * 2)
'        .ForeColor = &HFA8B10       'DCM_LightBlue   '&H864B24
'        .Appearance = aPlate
'        .BorderStyle = bsNone
'        .YHeight = lblPrgBar.Height - 10 ' 260
        DoEvents
    End With
    
    If optPrint(0).Value And (Not gUsingInWardMenu) Then
        If optBussDiv(0).Value Then
            Open App.Path & "\LIS_REPORT_" & Format(Now, CS_DateDbFormat) & "_�ܷ�.log" For Append As lngFileNo
        ElseIf optBussDiv(1).Value Then
            Open App.Path & "\LIS_REPORT_" & Format(Now, CS_DateDbFormat) & "_����.log" For Append As lngFileNo
        Else
            Open App.Path & "\LIS_REPORT_" & Format(Now, CS_DateDbFormat) & "_����.log" For Append As lngFileNo
        End If
    End If
    
'    Dim objWard As clsBasisData
    Dim strWard As String
    
    With tblOrder
        
        For i = 1 To .MaxRows
            
On Error GoTo Err_Trap1

            objProgress.Value = i
            
            .Row = i
            
            .Col = 1
            If .Value = 0 Then
                
                .TopRow = i
                
                .Col = 5    'ȯ�ڸ�
                objProgress.Message = .Value & " ȯ���� ������� ����ϰ� �ֽ��ϴ�... ( " & i & " / " & .MaxRows & " )"

                .Col = 4    'ȯ��ID
                strPtId = .Value
                
                .Col = 15   '���ڼ��� Path
                strImgPath = .Value

                .Col = 16   '���� ����
                strTestDiv = .Value
                
                picESign.Picture = LoadPicture(strImgPath)

                Set objReport = New clsBatchReport

                'Dictionary�� ���..����Ʈ ���
                .Col = 2:
'                Set objWard = Nothing
'                Set objWard = New clsBasisData
                strWard = GetWardNm(medGetP(.Value, 1, "-"))
'                Set objWard = Nothing
                
                If strWard <> "" Then
                    objReport.Ward = strWard
                    
                    If objReport.Ward <> "" Then
                        objReport.Ward = objReport.Ward & " " & Mid(.Value, Len(medGetP(.Value, 1, "-")) + 2)
                    Else
                        objReport.Ward = Mid(.Value, Len(medGetP(.Value, 1, "-")) + 2)
                    End If
                End If
                
'                If ObjLISComCode.WardID.Exists(medgetp(.Value, 1, "-")) = True Then
'                    ObjLISComCode.WardID.KeyChange (medgetp(.Value, 1, "-"))
'                    objReport.Ward = ObjLISComCode.WardID.Fields("wardnm")
'
'                    If objReport.Ward <> "" Then
'                        objReport.Ward = objReport.Ward & " " & Mid(.Value, Len(medgetp(.Value, 1, "-")) + 2)
'                    Else
'                        objReport.Ward = Mid(.Value, Len(medgetp(.Value, 1, "-")) + 2)
'                    End If
'                End If
                
                .Col = 3:  objReport.Doct = .Value
                .Col = 4:  objReport.ptid = .Value
                .Col = 5:  objReport.PtNm = .Value
                .Col = 6:  objReport.PtSex = medGetP(.Value, 1, "/")
                           objReport.PtAge = medGetP(.Value, 2, "/")
                .Col = 10: objReport.VfyDt = .Value
                '.Col = 11: objReport.VfyDt = objReport.VfyDt & " " & .Value
                .Col = 12: objReport.VfyNM = .Value
                .Col = 13: objReport.MdfDt = .Value         '������
                .Col = 17: objReport.ICD = .Value
                
                
                
                '�������� ����Ҷ��� ����Ʈ���� �����/ȸ���� ǥ��
                If gUsingInWardMenu Then
                    objReport.Rouding = optPrint(3).Value       'ȸ������Ʈ ����
                    objReport.Reprint = optPrint(2).Value       '����� ����
                    objReport.BatchReprint = True
                Else
                    objReport.Rouding = optPrint(3).Value       'ȸ������Ʈ ����
                    objReport.Reprint = optPrint(2).Value       '����� ����
                    objReport.BatchReprint = optPrint(1).Value
                End If
                objReport.Special = IIf(strTestDiv = enTestDiv.TST_SpeTest, True, False)
                
                .Col = 18:
'                Set objWard = Nothing
'                Set objWard = New clsBasisData
                strWard = GetDeptNm(.Value)
'                Set objWard = Nothing
                
                If strWard <> "" Then
                    objReport.Dept = .Value
                    objReport.DeptNm = strWard
                End If
'                If ObjLISComCode.DeptCd.Exists(.Value) Then
'                    Call ObjLISComCode.DeptCd.KeyChange(.Value)
'                    objReport.Dept = .Value
'                    objReport.DeptNm = ObjLISComCode.DeptCd.Fields("deptnm")
'                End If
                
                If optPrint(0).Value And (Not gUsingInWardMenu) Then
                    Print #lngFileNo, "( " & i & " / " & .MaxRows & " )  " & Now & "   " & strPtId & "," & objReport.PtNm & "," & objReport.DeptNm & "," & objReport.Ward
                End If
                
                objReport.RiPrint = "������"
                Call objReport.ReportForOnePatient(strPtId, strStartDate, Format(dtpVfyDt.Value, CS_DateDbFormat), _
                                                   strTestDiv, strImgPath, picESign, objProgress, strLastDt, strLastTm)
            End If
        Next
        objReport.RiPrint = ""
    End With

    If optPrint(0).Value And (Not gUsingInWardMenu) Then
        Close #lngFileNo
    End If

    MouseDefault
    
    Set objProgress = Nothing
    Set objProgress1 = Nothing
    
    If lngErrCount > 0 Then
        For i = tblOrder.DataRowCnt To 1 Step -1
            tblOrder.Row = i
            tblOrder.Col = 20
            If tblOrder.Value = "0" Then
                tblOrder.Action = ActionDeleteRow
            End If
        Next
        MsgBox "���� ȯ�ڵ��� ����� ��� �� ������ �߻��߽��ϴ�. �ٽ� ����Ͻʽÿ�.", vbExclamation, "����"
    Else
        cmdClear_Click
    End If
    
    Exit Sub
    
Err_Trap:
'==================================================================
    'DBConn.RollbackTrans
'==================================================================
    If optPrint(0).Value And (Not gUsingInWardMenu) Then
        Print #lngFileNo, "DB ERROR : " & Err.Description
    End If
On Error GoTo Err_Trap
    Resume Next

Err_Trap1:
    If optPrint(0).Value And (Not gUsingInWardMenu) Then
        Print #lngFileNo, "VB ERROR : " & Err.Description
    End If
On Error GoTo Err_Trap1
    Resume Next
    
End Sub

Private Sub cmdQuery_Click()

    Dim objReport   As New clsBatchReport
    Dim objESign    As clsLISElectronSign
    Dim objProgress As clsProgress
    Dim Rs          As New Recordset
    Dim rs1         As New Recordset
    Dim EmpRs       As Recordset
    Dim strWA       As String
    Dim strTable    As String
    Dim strWorkArea As String
    Dim strAccDt    As String
    Dim strAccseq   As String
    Dim strReferral As String
    Dim strSex      As String
    Dim strStsCd    As String
    Dim strMsg      As String
    Dim i           As Long
    Dim strEmpId    As String
    Dim strBussDiv  As String
    Dim strChkLoad  As String
    Dim strTestDiv  As String
    Dim strKey      As String
    Dim strDOB      As String
    Dim strSQL      As String
    'Dim strSEX      As String
    
    tblOrder.MaxRows = 0
    lblCnt.Caption = 0
    
    strBussDiv = IIf(optBussDiv(0).Value, enBussDiv.BussDiv_OutPatient, enBussDiv.BussDiv_InPatient)
    
    If optPrint(2).Value = True Then
        strStartDate = Format(dtpVfyDt.Value, CS_DateDbFormat)
    Else
        strStartDate = Format(DateAdd("d", -2, dtpVfyDt.Value), CS_DateDbFormat)
    End If
    
    '���α׷����� ����..
    Set objProgress = New clsProgress
    objProgress.Container = MainFrm.stsbar
    objProgress.Message = "�ڷḦ �а� �ֽ��ϴ�..."
    objProgress.Max = 100
'    objProgress.Caption = "ó�����Դϴ�."
'    objProgress.Mode = 0
'    objprogress.message = "�ڷḦ �а� �ֽ��ϴ�."
'    objProgress.Max = 100
'    objProgress.Min = 0
'    objProgress.Value = 0
'    objProgress.Visible = True
    
    objSql.RiPrint = "������"

    
    Dim strWard As String
    Dim strDoct As String
    
    If txtWardId.Text <> CS_AllCaption Then strWard = txtWardId.Text
    If txtDoctId.Text <> CS_AllCaption Then strDoct = txtDoctId.Text
    
    If optPrint(2).Value = True Then
        '���������
        Rs.Open objSql.GetAccLAbNoLIS201(txtPtId.Text, Format(dtpVfyDt.Value, CS_DateDbFormat)), DBConn
        tblOrder.ZOrder 0
        
    ElseIf optPrint(0).Value = True Then
        '�ϰ����
        Rs.Open objSql.RiReportList(strStartDate, Format(dtpVfyDt.Value, CS_DateDbFormat), strBussDiv, "", strWard, strDoct), DBConn
        tblOrder.ZOrder 0
    ElseIf optPrint(1).Value = True Then
        '�ϰ������
        Rs.Open objSql.RiReportList(strStartDate, Format(dtpVfyDt.Value, CS_DateDbFormat), strBussDiv, "Y", strWard, strDoct), DBConn
        tblOrder.ZOrder 0
    End If
    
    If Rs.EOF Then
        Set objProgress = Nothing
        MsgBox "�ش� ����Ÿ�� �����ϴ�.", vbInformation, "����� ���"
        GoTo Nodata
    End If
    
'    Dim objEmp As clsBasisData
    Dim strEmp As String
    
    strKey = ""
    With tblOrder
        
        If Rs.RecordCount > 0 Then

            '���α׷����� ����..
            objProgress.Max = Rs.RecordCount
            objProgress.Min = 0
            objProgress.Value = 0

            .ReDraw = False
            
            i = 1
            Do Until Rs.EOF = True
            
                If strKey = "" & Rs.Fields("deptcd").Value & _
                                 Rs.Fields("ptid").Value & _
                                 Rs.Fields("testdiv").Value Then
                    'ȯ��/�����/���������� ���� ��쿣 �������ο� �����ϸ� �����ֱ�...
                    If "" & Rs.Fields("stscd").Value = enStsCd.StsCd_LIS_Modify Then
                        .Col = 9
                        .Value = "����"
                    End If
                    If Trim("" & Rs.Fields("mfydt").Value) <> "" Then
                        .Col = 13
                        .Value = Format(Mid("" & Rs.Fields("mfydt").Value, 3), CS_DateShortMask)
                    End If
                    GoTo Skip
                End If
                    
                .MaxRows = i
                .Row = i

                .Col = 2: .Value = "" & Rs.Fields("location").Value
                .Col = 18: .Value = "" & Rs.Fields("deptcd").Value
                
                If optPrint(2).Value Then
                    If lblWard.Caption <> "" Then
                        .Col = 2: .Value = lblWard.Caption
                        .Col = 19: .Value = lblWard.Caption
                    End If
                Else
                    If optBussDiv(1).Value Then
                        .Col = 19: .Value = "" & Rs.Fields("location").Value
                    End If
                End If
'                Set objEmp = Nothing
'                Set objEmp = New clsBasisData
                strEmp = GetEmpNm(Rs.Fields("majdoct").Value & "")
'                Set objEmp = Nothing
                
                .Col = 3: .Value = strEmp 'GetEmpName(rs.Fields("majdoct").Value & "")
                .Col = 4: .Value = "" & Rs.Fields("ptid").Value
                .Col = 5: .Value = "" & Rs.Fields("ptnm").Value
                

                If IsNumeric("" & Rs.Fields("sex").Value) Then
                    strSex = IIf(Val("" & Rs.Fields("sex").Value) Mod 2 = 1, "��", "��")
                Else
                    strSex = IIf("" & Rs.Fields("sex").Value = "M", "��", "��")
                End If
                
                .Col = 6: .Value = strSex '
                          
                           strDOB = Rs.Fields("dob").Value & ""
                           If Len(strDOB) = 6 Then strDOB = strDOB & "01"
                            .Value = .Value & "/" & DateDiff("yyyy", Format(strDOB, CS_DateMask), Now)
                
                .Col = 16
                .Value = "" & Rs.Fields("testdiv").Value
                
                .Col = 7
                Select Case "" & Rs.Fields("testdiv").Value
                    Case enTestDiv.TST_RouTest
                        .Value = "�Ϲ�"
                    Case enTestDiv.TST_SpeTest
                        .Value = "��Ÿ"
                    Case enTestDiv.TST_MicTest
                        .Value = "�̻���"
                End Select

                .Col = 8: .Value = 1
                .Col = 9
                Select Case "" & Rs.Fields("stscd").Value
                Case enStsCd.StsCd_LIS_MidRst
                    .Value = "�߰�"
                Case enStsCd.StsCd_LIS_FinRst
                    .Value = "����"
                Case enStsCd.StsCd_LIS_Modify
                    .Value = "����"
                End Select

                .Col = 10: .Value = Format(Mid("" & Rs.Fields("vfydt").Value, 3), CS_DateShortMask)
                .Col = 11: .Value = Format(Mid("" & Rs.Fields("vfytm").Value, 1, 4), CS_TimeShortMask)

                strEmpId = "" & Rs.Fields("vfyid").Value
                .Col = 14: .Value = strEmpId
'                Set objEmp = Nothing
'                Set objEmp = New clsBasisData
                strEmp = GetEmpNm(Rs.Fields("majdoct").Value & "")
'                Set objEmp = Nothing
                
                .Col = 12:  .Value = strEmp ' GetEmpName(strEmpId)
               


                .Col = 13: .Value = Format(Mid("" & Rs.Fields("mfydt").Value, 3), CS_DateShortMask)
                '�ӻ�����....
                Dim objDisease  As New clsDisease
                
                objDisease.ptid = Rs.Fields("ptid").Value
                
                .Col = 17: .Value = objDisease.Disease
                
                Set objDisease = Nothing
                
                strKey = "" & Rs.Fields("deptcd").Value & _
                              Rs.Fields("ptid").Value & _
                              Rs.Fields("testdiv").Value
                              
                i = i + 1
                objProgress.Value = objProgress.Value + 1
Skip:
                Rs.MoveNext
            Loop
            Set objProgress = Nothing
            .ReDraw = True
            lblCnt.Caption = .MaxRows
        Else
            If optPrint(0).Value = True Then
                strMsg = "�������"
            ElseIf optPrint(1).Value = True Then
                strMsg = "�ϰ������"
            ElseIf optPrint(2).Value = True Then
                strMsg = "���������"
            End If
            MsgBox strMsg & " ������ �����ϴ�.", vbCritical, "��� ���"
            medClearTable tblOrder
            tblOrder.MaxRows = 0
            lblCnt.Caption = 0
        End If

    End With

Nodata:
    Set Rs = Nothing
    Set objProgress = Nothing

End Sub
Private Sub TxtClear()
    '����� ��� ����
    dtpVfyDt.Value = GetSystemDate

    '����� ��¿�������Ʈ
    medClearTable tblOrder
    
    With tblList
        .Row = 0: .Row2 = .MaxRows
        .Col = 2: .Col2 = .MaxCols
        .BlockMode = True
        .Text = ""
        .BlockMode = False
    End With
    
    lblWard.Caption = ""
    tblOrder.MaxRows = 0
    tblOrdSheet.MaxRows = 0
    tblOrder.ZOrder 0

    txtWardId.Text = "(��ü)"
    txtDoctId.Text = "(��ü)"
    lblCnt.Caption = 0
    chkAllWard.Value = 1
    chkAllDoct.Value = 1
    txtPtId.Text = ""
    lblPtNm.Caption = ""
    lblSexAge.Caption = ""

    cmdPreview.Caption = "�̸�����(&V)"
    cmdPreview.Tag = ""
    tblOrdSheet.Visible = False
    tblOrdSheet.ZOrder 1

End Sub
Private Sub chkAllDoct_Click()

    lblDoctNm.Caption = ""
    txtDoctId.Text = Choose(chkAllDoct.Value + 1, "", CS_AllCaption)
    txtDoctId.Enabled = Choose(chkAllDoct.Value + 1, True, False)
    cmdDoctList.Enabled = Choose(chkAllDoct.Value + 1, True, False)
End Sub

Private Sub chkAllWard_Click()

    lblWardNm.Caption = ""
    txtWardId.Text = Choose(chkAllWard.Value + 1, "", CS_AllCaption)
    txtWardId.Enabled = Choose(chkAllWard.Value + 1, True, False)
    cmdWardList.Enabled = Choose(chkAllWard.Value + 1, True, False)
End Sub
Private Sub Form_Load()
    lblWardNm.Caption = ""
    lblWard.Caption = ""
    
    optBussDiv(0).Enabled = True
    optBussDiv(0).Value = True

    blnLoadChk = False
    TxtClear

End Sub
Private Sub cmdClear_Click()
    TxtClear
End Sub
Private Sub cmdPreview_Click()
    
    Dim i As Long
    Dim strPtId    As String
    Dim strPtNm    As String
    Dim strVfyDt   As String
    Dim strTestDiv As String
    
    If cmdPreview.Tag = "1" Then
        cmdPreview.Caption = "�̸�����(&V)"
        cmdPreview.Tag = ""
        tblOrdSheet.Visible = False
        tblOrdSheet.ZOrder 1
    Else
        If tblOrder.MaxRows = 0 Then Exit Sub
        
        Dim objProgress As New clsProgress
        objProgress.Container = MainFrm.stsbar
        objProgress.Message = "�ڷḦ �а� �ֽ��ϴ�..."
        objProgress.Max = tblOrder.MaxRows
'        objProgress.Caption = "ó�����Դϴ�."
'        objProgress.Mode = 0
'        objprogress.message = "�ڷḦ �а� �ֽ��ϴ�."
'        objProgress.Max = tblOrder.MaxRows
'        objProgress.Min = 0
'        objProgress.Value = 0
'        objProgress.Visible = True
        
        tblOrdSheet.MaxRows = 0
        
        For i = 1 To tblOrder.MaxRows
            objProgress.Value = i
            tblOrder.Row = i
            tblOrder.Col = 1
            If tblOrder.Value = 1 Then GoTo Skip
            tblOrder.Col = 4:  strPtId = tblOrder.Value
            tblOrder.Col = 5:  strPtNm = tblOrder.Value
            tblOrder.Col = 16: strTestDiv = tblOrder.Value
            strVfyDt = Format(dtpVfyDt.Value, CS_DateDbFormat)
            objProgress.Message = strPtNm & "ȯ���� ��������� �а� �ֽ��ϴ�."
            DoEvents
            Call DisplayOrders(strPtId, strPtNm, strVfyDt, strTestDiv)
Skip:
        Next
        cmdPreview.Caption = "�ݱ�(&B)"
        cmdPreview.Tag = "1"
        tblOrdSheet.Visible = True
        tblOrdSheet.ZOrder 0
    End If
End Sub

Private Sub cmdWardList_Click()
'% �����ڵ� ����Ʈ�� �˾��Ѵ�.

    Dim objMyList As New clsPopUpList
'    Dim objWard As New clsBasisData
    Dim strCaption As String
    Dim strHead As String
    
    
    If optBussDiv(0).Value Then
        strCaption = "����� ��ȸ"
        strHead = "�μ��ڵ�;�μ���"
    Else
        strCaption = "���� ��ȸ"
        strHead = "�����ڵ�;������"
    End If
    
    With objMyList

        .FormCaption = strCaption
        .ColumnHeaderText = strHead
        .Tag = "WardID"
        Me.ScaleMode = 1
        If optBussDiv(0).Value Then
'            Call .ListPop(, 3950, 6300, ObjLISComCode.DeptCd)
            Call .LoadPopUp(GetSQLDeptList) ', 3950, 6300)
        Else
'            Call .ListPop(, 3950, 6300, ObjLISComCode.WardID)
            Call .LoadPopUp(GetSQLWardList) ', 3950, 6300)
        End If
        
        txtWardId.Text = medGetP(.SelectedString, 1, ";")
        lblWardNm.Caption = medGetP(.SelectedString, 2, ";")

    End With
    
'    Set objWard = Nothing
    Set objMyList = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objSql = Nothing
End Sub
Private Sub cmdDoctList_Click()

'% ��ġ�� ����Ʈ�� �˾��Ѵ�.

    Dim objMyList As New clsPopUpList
'    Dim objDoct As New clsBasisData

    With objMyList
        .FormCaption = "��ġ�Ǹ���Ʈ"

        .ColumnHeaderText = "�ǻ�ID;�ǻ��"
        .Tag = "DoctID"
        Me.ScaleMode = 1
'        Call .ListPop(getdoctlistsql, 3950, 6300)
        Call .LoadPopUp(GetSQLDoctList) ', 3950, 6300)

        txtDoctId.Text = medGetP(.SelectedString, 1, ";")
        lblDoctNm.Caption = medGetP(.SelectedString, 2, ";")

    End With
    
'    Set objDoct = Nothing
    Set objMyList = Nothing
End Sub

Private Sub cmdExit_Click()
    Set objSql = Nothing
    Unload Me
    
    RaiseEvent FormClose
End Sub

Private Function DisplayOrders(ByVal pPtId As String, ByVal pPtNm As String, ByVal pVfyDt As String, ByVal pTestDiv As String) As Boolean

    Dim SqlStmt         As String
    Dim ColCnt          As Integer
    Dim tmpTestNm       As String
    
    Dim SvKeyDt         As String
    Dim SvSpcNm         As String
    Dim pWorkArea       As String
    Dim pAccDt          As String
    Dim pAccSeq         As String
    Dim strKeyFld       As String
    Dim strNotice       As String
    Dim strTmp          As String
    Dim i               As Integer
    Dim j               As Integer
    Dim MySql           As New clsLISSqlReview     'Sql�� Ŭ����
    Dim tmpRs           As New Recordset
    Dim tVfyDt          As String
    
    
    Me.Enabled = False
   
    MouseRunning
    tVfyDt = Format(DateAdd("d", -2, dtpVfyDt.Value), CS_DateDbFormat)
    MySql.RiPrint = "������"
    'ó����/������ ����
    SqlStmt = MySql.SqlQueryAllResults(pPtId, "examdt", tVfyDt, pVfyDt, pTestDiv)
    
    'Query
    tmpRs.Open SqlStmt, DBConn
    
    SvKeyDt = "": SvSpcNm = ""
    
    DoEvents
   
    ReDim aryMesg(0)
    DisplayOrders = False
    
    If tmpRs.EOF Then GoTo Nodata
    
    With tblOrdSheet
      
        '.ReDraw = False
      
        Do Until tmpRs.EOF
         
            If Trim("" & tmpRs.Fields("RstCd").Value) = "" Then GoTo Skip
            
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1: .Value = pPtId: .ForeColor = DCM_Gray
            .Col = 2: .Value = pPtNm:  .ForeColor = DCM_Gray
            
            If SvKeyDt <> Trim("" & tmpRs.Fields("KeyDate").Value) Then
                .Col = 3:   .Value = Trim("" & tmpRs.Fields("KeyDate").Value)
                            .FontBold = True: .ForeColor = vbBlack       '-- ������
                .Col = 4:   .Value = Trim("" & tmpRs.Fields("SpcNm").Value)
                            .FontBold = True: .ForeColor = DCM_LightRed  '-- ��ü��
                SvKeyDt = Trim("" & tmpRs.Fields("KeyDate").Value)
                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
                .Col = 1:   .FontBold = True: .ForeColor = vbBlack
                .Col = 2:   .FontBold = True: .ForeColor = vbBlack
            Else
                .Col = 3:   .Value = "":
                            .FontBold = True: .ForeColor = vbBlack       '-- ó����
                            If SvSpcNm <> Trim("" & tmpRs.Fields("SpcNm").Value) Then
                                .Col = 4:
                                .Value = Trim("" & tmpRs.Fields("SpcNm").Value)
                                .FontBold = True: .ForeColor = DCM_LightRed  '-- ��ü��
                                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
                            Else
                                .Col = 4:
                                .Value = "":
                                .FontBold = True: .ForeColor = DCM_LightRed  '-- ��ü��
                            End If
            End If
            
            .Col = 34:  .Value = Trim("" & tmpRs.Fields("KeyDate").Value)   'ó����
            .Col = 35:  .Value = Trim("" & tmpRs.Fields("SpcNm").Value)     '��ü��
            
            .Col = 5:   '-- �˻��
                        .ForeColor = DCM_MidBlue
                        tmpTestNm = Mid(Trim("" & tmpRs.Fields("TestLongNm").Value), 1, 33)
                        If (Trim("" & tmpRs.Fields("DetailFg").Value) = "" And _
                            Trim("" & tmpRs.Fields("DetailItem").Value) = "") Or _
                            Trim("" & tmpRs.Fields("RstDiv").Value) = "*" Then
                            
                            .Value = tmpTestNm & " " & String(35 - Len(tmpTestNm), ".")
                        Else
                            .Value = Space(4) & tmpTestNm & " " & String(35 - Len("  " & tmpTestNm), ".")
                        End If
                        
            .Col = 6:   '-- �����(�ڵ��� ���..)
                        .ForeColor = DCM_Brown   '����
                        If Trim("" & tmpRs.Fields("VfyDt").Value) = "" Then
                            .Value = "��Ȯ"
                            .ForeColor = DCM_MidGray: .FontBold = False:
                        Else
                            If Trim("" & tmpRs.Fields("RstCdNm").Value) = "" Then
                                .TypeHAlign = TypeHAlignCenter
                                .Value = Trim("" & tmpRs.Fields("RstCd").Value)
                            Else
                                .CellType = CellTypeEdit
                                .TypeHAlign = TypeHAlignLeft
                                .Value = " " & Trim("" & tmpRs.Fields("RstCdNm").Value)
                            End If
                            If Trim("" & tmpRs.Fields("SenFg").Value) = "Y" Then
                                .Value = "Growth"
                            ElseIf Trim("" & tmpRs.Fields("RstCd").Value) = "" Then
                                .Value = Space(3)
                            End If
                        End If
                        
            .Col = 7:   '-- �������
                        .Value = Trim("" & tmpRs.Fields("RstUnit").Value)
            
            .Col = 8    '-- High / Low
                        .Value = ""
                        If Trim("" & tmpRs.Fields("VfyDt").Value) <> "" Then
                            If Trim("" & tmpRs.Fields("HLDiv").Value) = HLDIV_HIGH_CD Then .Value = HLDIV_HIGH_FG: .ForeColor = DCM_LightRed
                            If Trim("" & tmpRs.Fields("HLDiv").Value) = HLDIV_LOW_CD Then .Value = HLDIV_LOW_FG:  .ForeColor = DCM_LightBlue
                            If Trim("" & tmpRs.Fields("HLDiv").Value) = "*" Then .Value = "*": .ForeColor = vbRed
                        End If
            
            .Col = 9:   '-- Delta/Panic
                        .Value = Trim("" & tmpRs.Fields("DPDiv").Value): .ForeColor = vbRed
            
            .Col = 10:   '-- ����ġ
                        If Trim("" & tmpRs.Fields("RstDiv").Value) <> "*" And Trim("" & tmpRs.Fields("TestDiv").Value) < "4" Then .Value = CS_QuestionMark
            
            .Col = 11:   '-- More Result...
                        .Value = "": .ForeColor = DCM_LightBlue
                        If Trim("" & tmpRs.Fields("TxtFg").Value) > "0" Then .Value = CS_FingerMark
                        If Trim("" & tmpRs.Fields("TxtFg").Value) = "Y" Then .Value = CS_FingerMark
                        If Trim("" & tmpRs.Fields("SenFg").Value) = "Y" Then .Value = CS_FingerMark
                        If (Trim("" & tmpRs.Fields("DetailFg").Value) = "" And _
                            Trim("" & tmpRs.Fields("DetailItem").Value) = "") Or _
                            Trim("" & tmpRs.Fields("RstDiv").Value) = "*" Then
                            If Trim("" & tmpRs.Fields("FootNoteFg").Value) = "1" Then .Value = CS_FingerMark
                            If Trim("" & tmpRs.Fields("RmkCd").Value) <> "" Then .Value = CS_FingerMark
                        End If
                        If Trim("" & tmpRs.Fields("DcFg").Value) = "1" Then .Value = .Value & "*"
                        If Trim("" & tmpRs.Fields("TestDiv").Value) = "4" Then .Value = CS_FingerMark    '�غκ���
                        If Trim("" & tmpRs.Fields("TestDiv").Value) = "5" Then .Value = CS_FingerMark    '��������
         
            .Col = 12: .Value = Trim("" & tmpRs.Fields("OrdDate").Value)        '-- ó����
            .Col = 13: .Value = Trim("" & tmpRs.Fields("OrdNo").Value)          '-- ó���ȣ
            .Col = 14: .Value = Trim("" & tmpRs.Fields("OrdDoct").Value)        '-- ó����
            .Col = 15: .Value = Trim("" & tmpRs.Fields("ColDtTm").Value)        '-- ä���Ͻ�
            .Col = 16: .Value = Trim("" & tmpRs.Fields("ColId").Value)          '-- ä����
            .Col = 17: .Value = Trim("" & tmpRs.Fields("RcvDtTm").Value)        '-- �����Ͻ�
            .Col = 18: .Value = Trim("" & tmpRs.Fields("RcvId").Value)          '-- ������
            .Col = 19: .Value = Trim("" & tmpRs.Fields("WorkArea").Value):  pWorkArea = .Value  'WorkArea
            .Col = 20: .Value = Trim("" & tmpRs.Fields("AccDt").Value):     pAccDt = .Value     'AccDt
            .Col = 21: .Value = Trim("" & tmpRs.Fields("AccSeq").Value):    pAccSeq = .Value    'AccSeq
            .Col = 22: .Value = Trim("" & tmpRs.Fields("LastRst").Value)        '-- �ֱٰ��
            .Col = 23: .Value = Trim("" & tmpRs.Fields("LstVfyDtTm").Value)     '-- �ֱٰ���Ͻ�
            .Col = 24: .Value = Trim("" & tmpRs.Fields("LastVfyId").Value)      '-- �ֱٰ�� ������
            .Col = 25: .Value = Trim("" & tmpRs.Fields("VfyDtTm").Value)        '-- �����Ͻ�
            .Col = 26: .Value = Trim("" & tmpRs.Fields("VfyId").Value)          '-- ������
            .Col = 27: .Value = Trim("" & tmpRs.Fields("Sex").Value)            '-- Sex
            .Col = 28: .Value = Trim("" & tmpRs.Fields("AgeDay").Value)         '-- AgeDay
            .Col = 29: .Value = Trim("" & tmpRs.Fields("TestCd").Value)         '-- �˻��ڵ�
            .Col = 30: .Value = Trim("" & tmpRs.Fields("SpcCd").Value)          '-- ��ü�ڵ�
            .Col = 31: .Value = Trim("" & tmpRs.Fields("VfyDt").Value)          '-- ������
            .Col = 32: .Value = Trim("" & tmpRs.Fields("TestDiv").Value)        '-- �˻籸��
            .Col = 33: .Value = Trim("" & tmpRs.Fields("DeptCd").Value)         '-- �����
            .Col = 36: .Value = Trim("" & tmpRs.Fields("TxtFg").Value)          '-- �Ұ߰������
            .Col = 37: .Value = Trim("" & tmpRs.Fields("FootNoteFg").Value)     '-- Footnote ����
            .Col = 38: .Value = Trim("" & tmpRs.Fields("RmkCd").Value)          '-- Remark �ڵ�
            .Col = 39: .Value = Trim("" & tmpRs.Fields("SenFg").Value)          '-- ������ ����
            .Col = 40: .Value = Trim("" & tmpRs.Fields("OrdDiv").Value)         '-- ó�汸��
            .Col = 41: .Value = Trim("" & tmpRs.Fields("UnitQty").Value)        '-- ��������
            .Col = 42: .Value = Trim("" & tmpRs.Fields("ReqDt").Value)          '-- ����������
            .Col = 43: .Value = Trim("" & tmpRs.Fields("ReqTm").Value)          '-- ���������ð�
            .Col = 44: .Value = Trim("" & tmpRs.Fields("WardId").Value)         '-- ����
            .Col = 45: .Value = Trim("" & tmpRs.Fields("HosilId").Value)        '-- ȣ��
            .Col = 46: .Value = Trim("" & tmpRs.Fields("RoomId").Value)        '-- ȣ��
            
'            ReDim Preserve aryMesg(UBound(aryMesg) + 1)
'            aryMesg(UBound(aryMesg).value) = Trim("" & tmpRs.fields("Mesg").value)    '-- �����Remark
            If Trim("" & tmpRs.Fields("Notice").Value) <> "" Then
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Col = 5
                .TypeEditMultiLine = False
                .ForeColor = vbBlack
                .Value = "�� Clinical Notice "  '& vbCrLf & Trim("" & tmpRs.fields("Notice").value)
                .RowHeight(.MaxRows) = .MaxTextRowHeight(.MaxRows)
                strNotice = Trim("" & tmpRs.Fields("Notice").Value)
                strNotice = Replace(strNotice, vbCr, "")
                strTmp = medShift(strNotice, vbLf)
                While strTmp <> ""
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .Col = 5
                    .TypeEditMultiLine = False
                    .ForeColor = &H747474
                    .Value = strTmp
                    strTmp = medShift(strNotice, vbLf)
                Wend
            End If
      
         
            DisplayOrders = True
Skip:
            tmpRs.MoveNext
        Loop
      
        .Row = -1: .Col = 6: .Col2 = 5
        .BlockMode = True
        .AllowCellOverflow = True
        .BlockMode = False
      
        .RowHeight(-1) = 11.5
        .ReDraw = True
      
        'If chkRefVal.Value = 0 Then GoTo ExitPos
        GoTo ExitPos
      
        Dim tmpTestCd As String
        Dim tmpSpcCd As String
        Dim tmpVfyDt As String
        Dim tmpSex As String
        Dim tmpAgeDay As String
        Dim tmpRs1 As New Recordset
        Dim tmpRefFromVal As Double
        Dim tmpRefToVal As Double
        Dim tmpRefCd As String
      
        DoEvents
        For i = 1 To .MaxRows
            '����ġ �˻�
            .Row = i
            .Col = 10: If .Value <> CS_QuestionMark Then GoTo RefSkip
            
            .Col = 27:  tmpSex = Trim(.Value)
            .Col = 28:  tmpAgeDay = Trim(.Value)
            .Col = 29:  tmpTestCd = Trim(.Value)
            .Col = 30:  tmpSpcCd = Trim(.Value)
            .Col = 31:  tmpVfyDt = Trim(.Value)
                        If tmpVfyDt = "" Then tmpVfyDt = Format(Now, CS_DateDbFormat)
         
            SqlStmt = MySql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, "B", tmpAgeDay)
            Set tmpRs1 = Nothing
            tmpRs1.Open SqlStmt, DBConn
            
            If tmpRs1.EOF Then
                '"B"(Both)�� �ش��ϴ� ����ġ�� ���� ��� ȯ�ڼ����� �ش��ϴ� ����Ÿ �˻�
                '--> ���� Both�� ��ϵ�.
                SqlStmt = MySql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, tmpSex, tmpAgeDay)
                Set tmpRs1 = Nothing
                tmpRs1.Open SqlStmt, DBConn
            End If
            If tmpRs1.EOF Then
                tmpRefCd = Space(5)
            Else
                tmpRefFromVal = Val("" & tmpRs1.Fields("RefValFrom").Value)
                tmpRefToVal = Val("" & tmpRs1.Fields("RefValTo").Value)
                tmpRefCd = Trim("" & tmpRs1.Fields("RefCd").Value)
                If tmpRefFromVal <> 0 Or tmpRefToVal <> 0 Then _
                   tmpRefCd = tmpRefFromVal & "  -  " & tmpRefToVal
            End If
            Set tmpRs1 = Nothing
            For j = i To .MaxRows
                .Row = j
                .Col = 29   '����ġ
                If Trim(.Value) = tmpTestCd Then _
                    .Col = 10: .Value = tmpRefCd: .ForeColor = DCM_Green
            Next
         
            DoEvents

RefSkip:
        Next
      
ExitPos:
        'If .MaxRows < 20 Then .MaxRows = 20
      
    End With
   
Nodata:
    Me.Enabled = True
    MouseDefault
    DoEvents
    Set tmpRs = Nothing
    Set tmpRs1 = Nothing
   
End Function

Private Sub optBussDiv_Click(Index As Integer)
    cmdClear_Click
End Sub

Private Sub optPrint_Click(Index As Integer)
    
    If optPrint(2).Value = True Then

        fraLabNo.Visible = True
        fraSetWard.Visible = False
        txtPtId.Text = ""
        txtPtId.SetFocus
    Else
        fraLabNo.Visible = False
        fraSetWard.Visible = True
    End If
    
    If optPrint(0).Value = True Then
        If gUsingInWardMenu Then
            chkAllWard.Value = 0
            chkAllWard.Visible = False
        End If
    End If
    

    optBussDiv(0).Enabled = True
    If gUsingInWardMenu Then
        chkAllDoct.Enabled = True
        chkAllDoct.Value = 1
    Else
        chkAllDoct.Enabled = True
        chkAllWard.Enabled = True
        chkAllDoct.Value = 1
        chkAllWard.Value = 1
    End If
    
    tblList.Visible = False
    tblOrder.Visible = True
    tblOrder.ZOrder 0

    dtpVfyDt.Value = GetSystemDate
    lblPtNm.Caption = ""
    lblSexAge.Caption = ""

    '����� ��¿�������Ʈ
    medClearTable tblOrder

    lblCnt.Caption = 0

End Sub

Private Sub GetTestlist()
    Dim Rs As New Recordset
    Dim strTestNM As String
    Dim ii As Long
    Dim jj As Long
    
    Rs.Open objSql.GetTestReportList, DBConn
    If Rs.RecordCount > 0 Then
        ii = 0
        jj = 0
        strTestNM = ""
        With tblList
        
            .Row = 1: .Row2 = .MaxRows
            .Col = 1: .Col2 = 1
            .BlockMode = True
            .AllowCellOverflow = False
            .BlockMode = False
            
            .ReDraw = False
            .MaxRows = Rs.RecordCount + 1
            .Row = ii: .Col = 0
            .Value = "�˻��/��Ϲ�ȣ" & vbNewLine & "ȯ�ڸ�" & vbNewLine & "����/����"
            ii = 1
            Rs.MoveFirst
            Do Until Rs.EOF
                ii = ii + 1
                .Row = ii
                .Col = 0
                .RowHeight(ii) = 9.5
                If Rs.Fields("panelfg").Value & "" = "D" Then strTestNM = Rs.Fields("cdval1").Value & ""
                jj = Len(strTestNM)
                If strTestNM = Mid(Rs.Fields("cdval1").Value & "", 1, jj) And jj <> "0" And strTestNM <> Rs.Fields("cdval1").Value & "" Then
                    .Value = Space(4) & Rs.Fields("field1").Value & "": .TypeHAlign = TypeHAlignLeft
                Else
                    .Value = Space(1) & Rs.Fields("field1").Value & "": .TypeHAlign = TypeHAlignLeft
                End If
                .Col = 1
                .Value = Rs.Fields("cdval1").Value & "": .ForeColor = vbWhite
                Rs.MoveNext
            Loop
            .ReDraw = True
        End With
    End If
    Set Rs = Nothing
End Sub

Private Sub tblOrder_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    
    If MsgFg Then Exit Sub
    
    Dim lngButtonValue As Long
    Dim i As Long
    Dim strDept As String
    
    With tblOrder
        .Row = Row
        .Col = Col
        lngButtonValue = .Value
        If .Value = 1 Then
            lblCnt.Caption = Val(lblCnt.Caption) - 1
        Else
            lblCnt.Caption = Val(lblCnt.Caption) + 1
            Exit Sub
        End If
        
        .Col = 2
        strDept = medGetP(.Value, 1, "-")
        For i = 1 To tblOrder.DataRowCnt
            MsgFg = True
            .Row = i
            .Col = 2
            If strDept = medGetP(.Value, 1, "-") Then
                .Col = 1
                .Value = lngButtonValue
            End If
            MsgFg = False
        Next
    End With
End Sub

Private Sub txtPtId_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtPtId_LostFocus()
    
    Dim objPatient As clsPatient       'ȯ�� Ŭ����
    
    If Not gUsingInWardMenu Then

        If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
        If Screen.ActiveControl Is Nothing Then Exit Sub
        
        If Screen.ActiveControl.Name = cmdExit.Name Then Exit Sub
        If Screen.ActiveControl.Name = cmdClear.Name Then Exit Sub
    
    End If
    
    If MsgFg Then Exit Sub
      
    If txtPtId.Text = "" Then
        
        Exit Sub
    End If
    
    Set objPatient = New clsPatient
    If IsNumeric(txtPtId.Text) Then
        txtPtId.Text = Format(txtPtId.Text, P_PatientIdFormat)
    End If
    
    With objPatient
        If Trim(txtPtId.Text) <> "" And .GETPatient(txtPtId.Text) Then
            lblPtNm.Caption = .PtNm
            lblSexAge.Caption = .SEXNM & " / " & .Age & " " & .AGEDIV
            If .WardID = "" Then
                lblWard.Caption = ""
            Else
                lblWard.Caption = .WardID & "-" & .ROOMID
            End If
            PtFg = True
            ClearFg = False
        Else
            If Screen.ActiveControl.Name = cmdExit.Name Then Exit Sub
            MsgFg = True
            MsgBox "��ϵ��� ���� ȯ��ID�Դϴ�.. �ٽ� �Է��ϼ���..", vbInformation
            txtPtId.SetFocus
            MsgFg = False
            PtFg = False
            Set objPatient = Nothing
            Exit Sub
        End If
    End With
    
    Set objPatient = Nothing

    Exit Sub

End Sub
