VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FSJ0201 
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "�˻��׸� �߰� ����"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows �⺻��
   Begin Threed.SSPanel pnlbottom 
      Align           =   2  '�Ʒ� ����
      Height          =   525
      Left            =   0
      TabIndex        =   0
      Top             =   4590
      Width           =   5625
      _Version        =   65536
      _ExtentX        =   9922
      _ExtentY        =   926
      _StockProps     =   15
      ForeColor       =   16576
      BackColor       =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      Font3D          =   3
      Alignment       =   1
      Begin Threed.SSPanel pnlMsg 
         Height          =   390
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   5505
         _Version        =   65536
         _ExtentX        =   9710
         _ExtentY        =   688
         _StockProps     =   15
         ForeColor       =   8388608
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelOuter      =   1
      End
   End
   Begin Threed.SSPanel pnlmain 
      Align           =   1  '�� ����
      Height          =   4575
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5625
      _Version        =   65536
      _ExtentX        =   9922
      _ExtentY        =   8070
      _StockProps     =   15
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   2
      BevelInner      =   1
      Begin FPSpread.vaSpread SpdCode 
         Height          =   3810
         Left            =   60
         OleObjectBlob   =   "FSJ0201.frx":0000
         TabIndex        =   3
         Top             =   90
         Width           =   5490
      End
      Begin VB.TextBox txtCd 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   660
         TabIndex        =   4
         Top             =   4050
         Width           =   2655
      End
      Begin Threed.SSCommand CmdClk 
         Height          =   555
         Left            =   3420
         TabIndex        =   5
         Top             =   3930
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   979
         _StockProps     =   78
         Caption         =   "View"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand CmdEsc 
         Cancel          =   -1  'True
         Height          =   555
         Left            =   4500
         TabIndex        =   6
         Top             =   3930
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   979
         _StockProps     =   78
         Caption         =   "Esc"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�ڵ��"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   90
         TabIndex        =   7
         Top             =   4125
         Width           =   540
      End
   End
End
Attribute VB_Name = "FSJ0201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DisplayInit()
    txtCd = ""
    
    'SpreadBackColor Option
    iSpdBackColorOption = 3
    
    With SpdCode
        .BlockMode = True
        .Col = -1
        .Col2 = -1
        .Row = -1
        .Row2 = -1
        .BackColorStyle = BackColorStyleUnderGrid
        .BackColor = SpdBackcolor(iSpdBackColorOption)   'GBR
        .EditModePermanent = True
        .BlockMode = False
        
        .BlockMode = True
        .Col = 2
        .Col2 = 4
        .Row = -1
        .Row2 = -1
        .Lock = True
        .BlockMode = False
        
        If giCodeHlpMode = 2 Then
            .BlockMode = True
            .Col = 1
            .Col2 = 1
            .Row = -1
            .Row2 = -1
            .ColHidden = True
            .BlockMode = False
        End If
        
        .MaxRows = 0
        .MaxRows = 15
    End With
End Sub

Private Sub CmdClk_Click()
    Dim i%
    Dim vChk
    Dim vTestNm
    Dim vTestCd
    Dim vTestGbn
    
    MousePointer = 11
    
    Erase gCodeHlpTable
    giCodeHlpCnt = 0

    With SpdCode
        For i = 1 To .MaxRows
            If giCodeHlpMode = 1 Then
                Call .GetText(1, i, vChk)
                
                If vChk = "1" Then
            'FSJ0201���� ���ο� �׸��߰�
                    giCodeHlpCnt = giCodeHlpCnt + 1
                    
                    ReDim Preserve gCodeHlpTable(giCodeHlpCnt)
                    
                    Call .GetText(2, i, vTestNm)
                    Call .GetText(3, i, vTestCd)
                    Call .GetText(4, i, vTestGbn)
                    
                    gCodeHlpTable(giCodeHlpCnt).sCodeNm = CStr(vTestNm)
                    gCodeHlpTable(giCodeHlpCnt).sCode = CStr(vTestCd)
                    gCodeHlpTable(giCodeHlpCnt).sGbn = CStr(vTestGbn)
                End If
            End If
                        
            If giCodeHlpMode = 2 Then
                .Row = i
                .Col = 2
                 If .BackColor = RGB(255, 230, 230) Then
                    giCodeHlpCnt = giCodeHlpCnt + 1
                    
                    ReDim Preserve gCodeHlpTable(giCodeHlpCnt)
                    
                    Call .GetText(2, i, vTestNm)
                    Call .GetText(3, i, vTestCd)
                    Call .GetText(4, i, vTestGbn)
                    
                    gCodeHlpTable(giCodeHlpCnt).sCodeNm = CStr(vTestNm)
                    gCodeHlpTable(giCodeHlpCnt).sCode = CStr(vTestCd)
                    gCodeHlpTable(giCodeHlpCnt).sGbn = CStr(vTestGbn)
                    
                    Exit For
                 End If
            End If
        Next
        
        .Col = 1
        .Col2 = SpdCode.MaxCols
        .Row = 1
        .Row2 = SpdCode.MaxRows
        .BlockMode = True
        .Action = SS_ACTION_CLEAR_TEXT
        .BlockMode = False
    End With
    
    MousePointer = 0
    
    Unload Me
    
    
End Sub

Private Sub CmdEsc_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        CmdEsc_Click
    End If
End Sub

Private Sub Form_Load()
    Dim ret%
    Dim i%
    Dim j%
    
    Call DisplayInit
    
    Me.KeyPreview = True
    
    For i = 1 To giCodeHlpCnt
        For j = 1 To giCodeHlpCnt
            If Format$(i, "00000") = gCodeHlpTable(j).sSeq Then
                If i > 10 Then
                    SpdCode.MaxRows = i
                End If
                
                Call SpdCode.SetText(2, i, gCodeHlpTable(j).sCodeNm & "")
                Call SpdCode.SetText(3, i, gCodeHlpTable(j).sCode & "")
                Call SpdCode.SetText(4, i, gCodeHlpTable(j).sGbn & "")
            End If
        Next
    Next
    
    Erase gCodeHlpTable
    giCodeHlpCnt = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i%
    
    For i = 1 To giCodeHlpCnt
        With gCallObject
            If giCodeHlpMode = 1 Then
                 .Text = Right(gCodeHlpTable(i).sCode, 3) & ""
            End If
        End With
    Next
    
End Sub

Private Sub SpdCode_Click(ByVal Col As Long, ByVal Row As Long)
    Dim vCd
    Dim vCdNm
    Dim vGbn
    
    If Row = 0 Then
        Exit Sub
    End If
    
    Call spdReverse(SpdCode, -1, -1, Row, Row, RGB(255, 230, 230), iSpdBackColorOption)
    
    If Col <> 1 Then
        SpdCode.Col = 1
        SpdCode.Row = Row
        
        If SpdCode.Text = "" Or SpdCode.Text = "0" Then
            SpdCode.Text = "1"
        Else
            SpdCode.Text = ""
        End If
    End If
    
    Call SpdCode.GetText(2, Row, vCdNm)
        
    txtCd = CStr(vCdNm)
    
End Sub

Private Sub SpdCode_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row <> 0 Then
        Call CmdClk_Click
        Unload Me
    End If
End Sub
