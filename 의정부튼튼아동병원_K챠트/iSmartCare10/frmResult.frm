VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmResult 
   BackColor       =   &H00BF8B59&
   Caption         =   "�����ȸ"
   ClientHeight    =   8730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15705
   Icon            =   "frmResult.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   15705
   Begin VB.PictureBox picHeader 
      Align           =   1  '�� ����
      Appearance      =   0  '���
      BackColor       =   &H00AE8B59&
      BorderStyle     =   0  '����
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   0
      ScaleHeight     =   885
      ScaleWidth      =   15705
      TabIndex        =   0
      Top             =   0
      Width           =   15705
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   13230
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   1290
         TabIndex        =   2
         Top             =   120
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
         Format          =   136904705
         CurrentDate     =   40457
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   315
         Left            =   1290
         TabIndex        =   3
         Top             =   450
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
         Format          =   136904705
         CurrentDate     =   40457
      End
      Begin BHButton.BHImageButton cmdSearch 
         Height          =   555
         Left            =   3450
         TabIndex        =   8
         Top             =   180
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   979
         Caption         =   "�����ȸ"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmResult.frx":030A
         TransparentPicture=   "frmResult.frx":0464
         ButtonAttrib    =   2
         ButtonTransparent=   1
         ForeColor       =   16777215
         BackColor       =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdSave 
         Height          =   555
         Left            =   4920
         TabIndex        =   9
         Top             =   180
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   979
         Caption         =   "��������"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmResult.frx":2E56
         TransparentPicture=   "frmResult.frx":2FB0
         ButtonAttrib    =   2
         ButtonTransparent=   1
         ForeColor       =   16777215
         BackColor       =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdExcel 
         Height          =   555
         Left            =   6390
         TabIndex        =   10
         Top             =   180
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   979
         Caption         =   "��������"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmResult.frx":59A2
         TransparentPicture=   "frmResult.frx":5AFC
         ButtonAttrib    =   2
         ButtonTransparent=   1
         ForeColor       =   16777215
         BackColor       =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdDelete 
         Height          =   555
         Left            =   7860
         TabIndex        =   11
         Top             =   180
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   979
         Caption         =   "�������"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmResult.frx":84EE
         TransparentPicture=   "frmResult.frx":8648
         ButtonAttrib    =   2
         ButtonTransparent=   1
         ForeColor       =   16777215
         BackColor       =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClear 
         Height          =   555
         Left            =   9330
         TabIndex        =   12
         Top             =   180
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   979
         Caption         =   "ȭ������"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmResult.frx":B03A
         TransparentPicture=   "frmResult.frx":B194
         ButtonAttrib    =   2
         ButtonTransparent=   1
         ForeColor       =   16777215
         BackColor       =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClose 
         Height          =   555
         Left            =   10800
         TabIndex        =   7
         Top             =   180
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   979
         Caption         =   "�ݱ�"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmResult.frx":DB86
         TransparentPicture=   "frmResult.frx":DCE0
         ButtonAttrib    =   2
         ButtonTransparent=   1
         ForeColor       =   16777215
         BackColor       =   16777215
         ImgOutLineSize  =   3
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "����"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   0
         Left            =   2760
         TabIndex        =   6
         Top             =   210
         Width           =   360
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "��ȸ�Ⱓ :"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   300
         TabIndex        =   5
         Top             =   210
         Width           =   840
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   795
         Left            =   90
         Top             =   60
         Width           =   12315
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "����"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   2
         Left            =   2760
         TabIndex        =   4
         Top             =   540
         Width           =   360
      End
   End
   Begin FPSpread.vaSpread spdResult 
      Height          =   8835
      Left            =   90
      TabIndex        =   1
      Top             =   900
      Width           =   21195
      _Version        =   393216
      _ExtentX        =   37386
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
      MaxCols         =   22
      MaxRows         =   20
      OperationMode   =   2
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   16777215
      SpreadDesigner  =   "frmResult.frx":106D2
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
End
Attribute VB_Name = "frmResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnCheck    As Boolean

Private Sub cmdClear_Click()
    
    spdResult.MaxRows = 0
    
End Sub

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub

Private Sub cmdDelete_Click()
    Dim lRow As Long
    
    If MsgBox("������ ����� �����Ͻðڽ��ϱ�?", vbYesNo + vbCritical, "�������") = vbYes Then
        With spdResult
            For lRow = 1 To .DataRowCnt
                .Row = lRow
                .Col = 1
                If .Value = 1 Then
                          SQL = " DELETE From PATRESULT " & vbCrLf
                    SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "' " & vbCrLf
                    SQL = SQL & "   AND SAVESEQ = " & Trim(GetText(spdResult, lRow, colSAVESEQ))
                    SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdResult, lRow, colBARCODE)) & "' "
                        
                    If DBExec(AdoCn_Local, SQL) Then
                        '-- ����
                    End If
                    
                    spdResult.Row = lRow
                    spdResult.Col = 1
                    spdResult.Value = 0
                End If
            Next lRow
        End With
        
        Call cmdSearch_Click
        
    End If
    
End Sub

Private Sub cmdExcel_Click()

'    Call spdResult.ExportToExcel(App.PATH & "\" & Format(Now, "yyyy-mm-dd") & "_�������.xls", "�������", "Log.Text")
    
    Dim sFileName As String
            
On Error GoTo ErrHandler

    If spdResult.DataRowCnt < 1 Then
        MsgBox "������ �ڷᰡ �����ϴ�.", , "�� ��"
        Exit Sub
    Else
        With CommonDialog1
            .CancelError = True
            .Flags = cdlOFNHideReadOnly
            .InitDir = App.PATH
            .Filter = "Excel Files (*.xls)|*.xls|All Files (*.*)|*.*"
            .Filename = App.PATH & "\" & Format(Now, "yyyy-mm-dd") & "_�������.xls"
            .ShowSave
            sFileName = CommonDialog1.Filename
            SaveExcel sFileName, spdResult
            MsgBox "���� ����Ϸ�", vbOKOnly + vbInformation, Me.Caption
        End With
    End If

Exit Sub
  
ErrHandler:
      
    ' ����ڰ� [���] ���߸� �������ϴ�.
    Exit Sub

End Sub

Private Sub cmdSave_Click()
    Dim lRow As Long
    Dim Res  As Integer
    
    If MsgBox("������ ����� �������Ͻðڽ��ϱ�?", vbYesNo + vbCritical, "�������") = vbYes Then
        With spdResult
            For lRow = 1 To .DataRowCnt
                .Row = lRow
                .Col = 1
                If .Value = 1 Then
                    
                    Res = SaveTransData(lRow, spdResult)
                
                    If Res = -1 Then
                        SetForeColor spdResult, lRow, lRow, 1, colSTATE, 255, 0, 0
                        SetText spdResult, "�������", lRow, colSTATE
                              
                              SQL = " UPDATE PATRESULT SET " & vbCrLf
                        SQL = SQL & "     SENDFLAG  = '1' " & vbCrLf
                        SQL = SQL & "   , SENDDATE  = '" & Format(Now, "yyyy-mm-dd") & "' " & vbCrLf
                        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "' " & vbCrLf
                        SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdResult, lRow, colBARCODE)) & "' "
                        
                        If DBExec(AdoCn_Local, SQL) Then
                            '-- ����
                        End If
                    Else
                        SetBackColor spdResult, lRow, lRow, 1, colSTATE, 202, 255, 112
                        SetText spdResult, "����Ϸ�", lRow, colSTATE
                        
                              SQL = " UPDATE PATRESULT SET " & vbCrLf
                        SQL = SQL & "     SENDFLAG  = '2' " & vbCrLf
                        SQL = SQL & "   , SENDDATE  = '" & Format(Now, "yyyy-mm-dd") & "' " & vbCrLf
                        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "' " & vbCrLf
                        SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdResult, lRow, colBARCODE)) & "' "
                        
                        If DBExec(AdoCn_Local, SQL) Then
                            '-- ����
                        End If
                        
                    End If
                    spdResult.Row = lRow
                    spdResult.Col = 1
                    spdResult.Value = 0
                End If
            Next lRow
        End With
    End If
    
End Sub

Private Sub cmdSearch_Click()

    spdResult.MaxRows = 0
    
    Call GetResultList(Format(dtpFrom.Value, "yyyy-mm-dd"), Format(dtpTo.Value, "yyyy-mm-dd"), spdResult)

End Sub

Private Sub Form_Load()
    
    dtpFrom.Value = Now
    dtpTo.Value = Now

    blnCheck = True
    
    spdResult.MaxRows = 0
    
    '-- �÷����̱⼳��
    Call SetColumnView(spdResult)
    
    '-- �˻�� ���̱�
    Call SetExamCode(spdResult)
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Resize()

    If Me.ScaleHeight = 0 Then Exit Sub

    spdResult.WIDTH = Me.ScaleWidth - 300
    spdResult.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - 300
    
End Sub

Private Sub spdResult_Click(ByVal Col As Long, ByVal Row As Long)
    Dim iRow As Long
    
    If Row = 0 And Col = colCHECKBOX Then
        With spdResult
            If blnCheck = False Then
                For iRow = 1 To .DataRowCnt
                    .Row = iRow
                    .Col = 1
                    
                    .Value = 1
                Next iRow
                blnCheck = True
            Else
                For iRow = 1 To .DataRowCnt
                    .Row = iRow
                    .Col = 1
                    
                    .Value = 0
                Next iRow
                blnCheck = False
            End If
        End With
    End If


    If Row > 0 And Col = colCHECKBOX Then
        If GetText(spdResult, Row, colCHECKBOX) = "1" Then
            Call SetText(spdResult, "0", Row, colCHECKBOX)
        Else
            Call SetText(spdResult, "1", Row, colCHECKBOX)
        End If
    End If


End Sub

Private Sub spdResult_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sRow        As Long
    Dim sCol        As Long
    Dim strNewBarNo As String
    
    sRow = spdResult.ActiveRow
    sCol = colPNAME
    
    If sRow = 0 Then
        Exit Sub
    End If
    
    strNewBarNo = GetText(spdResult, sRow, sCol)
    
    If KeyCode = vbKeyDelete Then
        
        If MsgBox(strNewBarNo & " �� ����ðڽ��ϱ�?", vbInformation + vbYesNo, "�˸�") = vbNo Then
            Exit Sub
        End If
        
        DeleteRow spdResult, sRow, sRow
        spdResult.MaxRows = spdResult.MaxRows - 1
    
    End If

End Sub
