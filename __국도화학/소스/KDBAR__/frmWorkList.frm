VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmWorkList 
   BackColor       =   &H00FFFFFF&
   Caption         =   "��ũ����Ʈ"
   ClientHeight    =   9645
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   16470
   Icon            =   "frmWorkList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   16470
   StartUpPosition =   1  '������ ���
   Begin VB.PictureBox picHeader 
      Align           =   1  '�� ����
      Appearance      =   0  '���
      BackColor       =   &H00800000&
      BorderStyle     =   0  '����
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   16470
      TabIndex        =   0
      Top             =   0
      Width           =   16470
      Begin VB.CommandButton cmdWorkPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ũ���"
         Height          =   375
         Left            =   8340
         Style           =   1  '�׷���
         TabIndex        =   14
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ũ��ȸ"
         Height          =   375
         Left            =   4920
         Style           =   1  '�׷���
         TabIndex        =   13
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmdSend 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ȭ������"
         Height          =   375
         Left            =   6060
         Style           =   1  '�׷���
         TabIndex        =   12
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�ݱ�"
         Height          =   375
         Left            =   9480
         Style           =   1  '�׷���
         TabIndex        =   11
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmdSendClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "����/�ݱ�"
         Height          =   375
         Left            =   7200
         Style           =   1  '�׷���
         TabIndex        =   10
         Top             =   180
         Width           =   1095
      End
      Begin VB.CheckBox chkSave 
         Appearance      =   0  '���
         BackColor       =   &H00800000&
         Caption         =   "��������"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   13380
         TabIndex        =   7
         Top             =   300
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdSeq 
         Caption         =   "Seq ��ġ"
         Height          =   375
         Left            =   14580
         TabIndex        =   6
         Top             =   210
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox txtSeqNo 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   12570
         TabIndex        =   4
         Text            =   "1"
         Top             =   180
         Width           =   645
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   1500
         TabIndex        =   1
         Top             =   180
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "���� ���"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   139001857
         CurrentDate     =   40457
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   315
         Left            =   3180
         TabIndex        =   2
         Top             =   180
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "���� ���"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   139001857
         CurrentDate     =   40457
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   465
         Left            =   4800
         Top             =   120
         Width           =   5865
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "��ȸ�Ⱓ :"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   270
         Width           =   930
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00C0FFFF&
         BorderWidth     =   2
         Height          =   465
         Left            =   270
         Top             =   120
         Width           =   4485
      End
      Begin VB.Label lblSeqNo 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "��ȣ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   11790
         TabIndex        =   5
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   0
         Left            =   2970
         TabIndex        =   3
         Top             =   270
         Width           =   150
      End
   End
   Begin FPSpread.vaSpread spdWork 
      Height          =   8835
      Left            =   30
      TabIndex        =   9
      Top             =   750
      Width           =   16395
      _Version        =   393216
      _ExtentX        =   28919
      _ExtentY        =   15584
      _StockProps     =   64
      ColsFrozen      =   22
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "���� ���"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   15921919
      GridShowVert    =   0   'False
      MaxCols         =   23
      MaxRows         =   20
      OperationMode   =   2
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   16777215
      SpreadDesigner  =   "frmWorkList.frx":06C2
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
End
Attribute VB_Name = "frmWorkList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub

Private Sub cmdSearch_Click()
    
    Call GetWorkList(Format(dtpFrom.Value, "yyyymmdd"), Format(dtpTo.Value, "yyyymmdd"), spdWork)
    
    spdWork.RowHeight(-1) = 15

End Sub

Private Sub cmdSend_Click()
    Dim i               As Integer
    Dim intRow          As Integer
    Dim intWRow         As Integer
    Dim intORow         As Integer
    Dim intWCol         As Integer
    Dim intOCol         As Integer
    Dim strBarno        As String
    Dim blnSame         As Boolean
    Dim varItems        As Variant
    Dim intItems        As Integer
    
    With spdWork
        For intWRow = 1 To .MaxRows
            .Row = intWRow
            .Col = colCHECKBOX
            If .Value = "1" Then
                blnSame = False
                strBarno = GetText(spdWork, intWRow, colBARCODE)
                For intORow = 1 To frmMain.spdOrder.MaxRows
                    frmMain.spdOrder.Row = intORow
                    frmMain.spdOrder.Col = colBARCODE
                    If strBarno = GetText(frmMain.spdOrder, intORow, colBARCODE) Then
                        blnSame = True
                    End If
                Next
                
                If blnSame = False Then
                    frmMain.spdOrder.MaxRows = frmMain.spdOrder.MaxRows + 1
                    intRow = frmMain.spdOrder.MaxRows
                    For i = colCHECKBOX To colSTATE
                        Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, i), intRow, i)
                
                        varItems = GetText(spdWork, intWRow, colITEMS)
                        varItems = Split(varItems, "/")
                        For intItems = 0 To UBound(varItems)
                            For intOCol = colSTATE + 1 To frmMain.spdOrder.MaxCols
                                frmMain.spdOrder.Row = 0
                                frmMain.spdOrder.Col = intOCol
                                If varItems(intItems) = Trim(frmMain.spdOrder.Text) Then
                                    .Row = frmMain.spdOrder.MaxRows
                                    Call SetText(frmMain.spdOrder, "��", frmMain.spdOrder.MaxRows, intOCol)
'                                    GoTo RST
                                End If
                            Next
                        Next
                    Next
                    
                    frmMain.spdOrder.RowHeight(-1) = 15
                End If
            End If
        Next
    End With
    
End Sub

Private Sub cmdSendClose_Click()
    
    Call cmdSend_Click
    
    Call cmdClose_Click
    
End Sub

Private Sub cmdSeq_Click()
    Dim intWRow         As Integer
    Dim intORow         As Integer
    Dim intWCol         As Integer
    Dim intOCol         As Integer
    Dim strBarno        As String
    Dim strSeq          As String
    Dim blnSame         As Boolean
    Dim varItems        As Variant
    Dim intItems        As Integer
    
    With spdWork
        For intWRow = 1 To .MaxRows
            .Row = intWRow
            .Col = colCHECKBOX
            If .Value = "1" Then
                For intORow = 1 To frmMain.spdOrder.MaxRows
                    If GetText(spdWork, intWRow, colSEQNO) = GetText(frmMain.spdOrder, intORow, colSEQNO) Then
                        
                        Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colBARCODE), intORow, colBARCODE)
                        DoEvents
                        If GetSampleInfo(intORow, frmMain.spdOrder) = -1 Then
                            'MsgBox "�Է��� ���ڵ忡�� ȯ�������� ã�� ���߽��ϴ�." & vbNewLine & " ���ڵ� ��ȣ�� Ȯ���ϼ���", vbOKOnly + vbCritical, Me.Caption
                        Else
                            '��������
                            SQL = ""
                            SQL = SQL & "UPDATE PATRESULT SET "
                            SQL = SQL & "  BARCODE       = '" & Trim(GetText(frmMain.spdOrder, intORow, colBARCODE)) & "'" & vbCr
                            SQL = SQL & " ,INOUT         = '" & Trim(GetText(frmMain.spdOrder, intORow, colINOUT)) & "'" & vbCr
                            SQL = SQL & " ,CHARTNO       = '" & Trim(GetText(frmMain.spdOrder, intORow, colCHARTNO)) & "'" & vbCr
                            SQL = SQL & " ,PID           = '" & Trim(GetText(frmMain.spdOrder, intORow, colPID)) & "'" & vbCr
                            SQL = SQL & " ,PNAME         = '" & Trim(GetText(frmMain.spdOrder, intORow, colPNAME)) & "'" & vbCr
                            SQL = SQL & " ,PSEX          = '" & Trim(GetText(frmMain.spdOrder, intORow, colPSEX)) & "'" & vbCr
                            SQL = SQL & " ,PAGE          = '" & Trim(GetText(frmMain.spdOrder, intORow, colPAGE)) & "'" & vbCr
''                            SQL = SQL & " ,PJUMIN        = '" & Trim(GetText(frmMain.spdOrder, intORow, colPJUMIN)) & "'" & vbCr
'                            SQL = SQL & " ,PANICVALUE    = '" & Trim(GetText(frmMain.spdOrder, intORow, colKEY1)) & "'" & vbCr
                            SQL = SQL & " WHERE EXAMDATE = '" & Trim(GetText(frmMain.spdOrder, intORow, colEXAMDATE)) & "'" & vbCr
                            SQL = SQL & "   AND SAVESEQ  = " & Trim(GetText(frmMain.spdOrder, intORow, colSAVESEQ)) & vbCr
                            SQL = SQL & "   AND EQUIPNO  = '" & gKUKDO.MACHCD & "' & vbCr"
                            
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ����
                            End If
                        End If
                        Exit For
                    End If
                Next intORow
            End If
        Next intWRow
    End With
End Sub

Private Sub cmdWorkPrint_Click()
    
    If spdWork.DataRowCnt < 1 Then
        MsgBox "����� �ڷᰡ �����ϴ�.", , "�� ��"
        Exit Sub
    Else
        spdWork.PrintOrientation = PrintOrientationPortrait
        spdWork.Action = 13
    End If
    

End Sub

Private Sub Form_Load()
    
    Call CtlInitializing

    '-- �÷����̱⼳��
    Call SetColumnView(spdWork)
    
    spdWork.ColWidth(spdWork.MaxCols) = 20
        
'    spdWork.MaxRows = 10
    
    
'    Dim i As Integer
'
'    For i = 1 To 10
'        Call SetText(spdWork, i, i, colBARCODE)
'        Call SetText(spdWork, i * 10, i, colITEMS)
'
'    Next
    '-- �˻�� ���̱�
'    Call SetExamCode(spdWork)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub CtlInitializing()
    
    spdWork.MaxRows = 0
    
    dtpFrom.Value = Now
    dtpTo.Value = Now
    
    txtSeqNo.Text = "1"
    
    '�������
    If gKUKDO.RSTTYPE = "1" Then
        lblSeqNo.Visible = True
        txtSeqNo.Visible = True
    Else
        lblSeqNo.Visible = False
        txtSeqNo.Visible = False
    End If
    
End Sub

Private Sub Form_Resize()
    
    If Me.ScaleHeight = 0 Then Exit Sub

    spdWork.WIDTH = Me.ScaleWidth - 300
    spdWork.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - 300

    spdWork.ColWidth(colSTATE + 1) = 60 '(spdWork.Width / 40) * intColSum

End Sub

Private Sub spdWork_Click(ByVal Col As Long, ByVal Row As Long)
    Dim i As Integer

    If Row = 0 And Col <> colCHECKBOX Then
        Call SetSpreadSort(spdWork, 0)
        Exit Sub
    End If
    
    If Row = 0 And Col = colCHECKBOX Then
        If GetText(spdWork, 1, colCHECKBOX) = "1" Then
            For i = 1 To spdWork.DataRowCnt
                Call SetText(spdWork, "0", i, colCHECKBOX)
            Next
        Else
            For i = 1 To spdWork.DataRowCnt
                Call SetText(spdWork, "1", i, colCHECKBOX)
            Next
        End If
    End If
    
    If Row > 0 And Col = colCHECKBOX Then
        If GetText(spdWork, Row, colCHECKBOX) = "1" Then
            Call SetText(spdWork, "0", Row, colCHECKBOX)
        Else
            Call SetText(spdWork, "1", Row, colCHECKBOX)
        End If
    End If
    
'    txtQuery.Visible = True
'    txtQuery.Text = GetText(spdWork, Row, colITEMS)
    
End Sub

Private Sub spdWork_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim i               As Integer
    Dim intRow          As Integer
    Dim intWRow         As Integer
    Dim intORow         As Integer
    Dim intWCol         As Integer
    Dim intOCol         As Integer
    Dim strBarno        As String
    Dim blnSame         As Boolean
    Dim varItems        As Variant
    Dim intItems        As Integer
    
    Dim strBarno_Work   As String
    
    If Row = 0 Then Exit Sub
    If Col <> colBARCODE Then
        Exit Sub
    End If
    
    intWRow = Row
    spdWork.Row = Row
    spdWork.Col = colBARCODE
    strBarno_Work = Trim(spdWork.Text)
    
    With frmMain.spdOrder
        blnSame = False
        For intORow = 1 To .MaxRows
            .Row = intORow
            .Col = colBARCODE
            If strBarno_Work = Trim(.Text) Then
                blnSame = True
                Exit For
            End If
        Next
        
        If blnSame = False Then
            frmMain.spdOrder.MaxRows = frmMain.spdOrder.MaxRows + 1
            intRow = frmMain.spdOrder.MaxRows
            
            For i = colCHECKBOX To colSTATE
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, i), intRow, i)

'                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colSPECNO), intRow, colSPECNO)
'                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colCHECKBOX), intRow, colCHECKBOX)
'                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colHOSPDATE), intRow, colHOSPDATE)
'                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colBARCODE), intRow, colBARCODE)
'                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colSEQNO), intRow, colSEQNO)
'                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colCHARTNO), intRow, colCHARTNO)
'                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPID), intRow, colPID)
'                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colINOUT), intRow, colINOUT)
'                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPNAME), intRow, colPNAME)
'                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPSEX), intRow, colPSEX)
'                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPAGE), intRow, colPAGE)
'    '            Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPJUMIN), introw, colPJUMIN)
'                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colOCNT), intRow, colOCNT)
                
                varItems = GetText(spdWork, intWRow, colITEMS)
                varItems = Split(varItems, "/")
                For intItems = 0 To UBound(varItems)
                    For intOCol = colSTATE + 1 To frmMain.spdOrder.MaxCols
                        frmMain.spdOrder.Row = 0
                        frmMain.spdOrder.Col = intOCol
                        If varItems(intItems) = Trim(frmMain.spdOrder.Text) Then
                            .Row = intRow
                            Call SetText(frmMain.spdOrder, "��", intRow, intOCol)
                        End If
                    Next
                Next
            Next
            
            frmMain.spdOrder.RowHeight(-1) = 15
        End If
    
    End With
    
End Sub

Private Sub spdWork_KeyPress(KeyAscii As Integer)
    Dim intRow As Integer
    Dim strSeq As String
    
    If KeyAscii = vbKeyReturn Then
        With spdWork
            If .ActiveCol = colSEQNO Then
                strSeq = GetText(spdWork, .ActiveRow, .ActiveCol)
                If Not IsNumeric(strSeq) Then
                    MsgBox "���ڸ� �Է��� �����մϴ�"
                    Exit Sub
                End If
                For intRow = .ActiveRow + 1 To .MaxRows
                    Call SetText(spdWork, strSeq + 1, intRow, colSEQNO)
                Next
            End If
        End With
    End If
    
End Sub
