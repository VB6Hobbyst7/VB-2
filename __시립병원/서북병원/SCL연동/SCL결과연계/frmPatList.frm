VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatList 
   Caption         =   "�׿��� �Ƿ�ȯ�� ����Ʈ"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10860
   BeginProperty Font 
      Name            =   "����ü"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11460
   ScaleWidth      =   18960
   Visible         =   0   'False
   WindowState     =   2  '�ִ�ȭ
   Begin FPSpread.vaSpread vasTemp1 
      Height          =   2115
      Left            =   8220
      TabIndex        =   20
      Top             =   4020
      Visible         =   0   'False
      Width           =   3585
      _Version        =   393216
      _ExtentX        =   6324
      _ExtentY        =   3731
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmPatList.frx":0000
   End
   Begin FPSpread.vaSpread vasOrder 
      Height          =   3765
      Left            =   1860
      TabIndex        =   19
      Top             =   4710
      Visible         =   0   'False
      Width           =   4425
      _Version        =   393216
      _ExtentX        =   7805
      _ExtentY        =   6641
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmPatList.frx":0228
   End
   Begin FPSpread.vaSpread vasPrint 
      Height          =   1215
      Left            =   1290
      TabIndex        =   16
      Top             =   3540
      Visible         =   0   'False
      Width           =   5385
      _Version        =   393216
      _ExtentX        =   9499
      _ExtentY        =   2143
      _StockProps     =   64
      ColHeaderDisplay=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   10
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmPatList.frx":0450
   End
   Begin FPSpread.vaSpread vasExcel 
      Height          =   1215
      Left            =   1410
      TabIndex        =   13
      Top             =   4980
      Visible         =   0   'False
      Width           =   5385
      _Version        =   393216
      _ExtentX        =   9499
      _ExtentY        =   2143
      _StockProps     =   64
      ColHeaderDisplay=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   10
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmPatList.frx":426D
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1980
      Top             =   3420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkAll 
      Height          =   285
      Left            =   660
      TabIndex        =   6
      Top             =   930
      Width           =   195
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   7905
      Left            =   60
      TabIndex        =   1
      Top             =   870
      Width           =   15135
      _Version        =   393216
      _ExtentX        =   26696
      _ExtentY        =   13944
      _StockProps     =   64
      ColHeaderDisplay=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   10
      MaxRows         =   499
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmPatList.frx":8046
   End
   Begin VB.Frame Frame1 
      Height          =   765
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   15135
      Begin �Ƿڰ������_����������.MDButton btnClose 
         Height          =   495
         Left            =   13710
         TabIndex        =   17
         Top             =   150
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "�ݱ�"
      End
      Begin VB.TextBox txtBarcode 
         Appearance      =   0  '���
         BackColor       =   &H00D8FEFE&
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1680
         TabIndex        =   15
         Top             =   225
         Width           =   1905
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "��ü��ȣ"
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   5100
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   2595
         Begin VB.OptionButton optStat 
            Caption         =   "���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   3480
            TabIndex        =   12
            Top             =   180
            Width           =   795
         End
         Begin VB.OptionButton optStat 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   2640
            TabIndex        =   11
            Top             =   180
            Width           =   795
         End
         Begin VB.OptionButton optStat 
            Caption         =   "�Է�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   2
            Left            =   1800
            TabIndex        =   10
            Top             =   120
            Width           =   765
         End
         Begin VB.OptionButton optStat 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   990
            TabIndex        =   9
            Top             =   120
            Value           =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton optStat 
            Caption         =   "���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   8
            Top             =   120
            Width           =   765
         End
      End
      Begin VB.ComboBox cboDate 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         ItemData        =   "frmPatList.frx":BDFA
         Left            =   120
         List            =   "frmPatList.frx":BE04
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   1515
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   47710209
         CurrentDate     =   39787
      End
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   315
         Left            =   3510
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   47710209
         CurrentDate     =   39787
      End
      Begin �Ƿڰ������_����������.MDButton btnExcel 
         Height          =   495
         Left            =   11250
         TabIndex        =   18
         Top             =   150
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Excel ����"
      End
      Begin �Ƿڰ������_����������.MDButton btnClear 
         Height          =   495
         Left            =   12570
         TabIndex        =   21
         Top             =   150
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Clear"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '����
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3270
         TabIndex        =   4
         Top             =   300
         Visible         =   0   'False
         Width           =   195
      End
   End
End
Attribute VB_Name = "frmPatList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnClear_Click()
    ClearSpread vasList
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnExcel_Click()
'ȭ�ϸ� ����ڰ� ���� �����ϵ��� ��

    Dim sFileName As String
    Dim lRow As Long
    Dim i As Integer
    Dim iCol As Integer
    
'On Error GoTo ErrHandler

    If MsgBox("Excel ������ �����Ͻðڽ��ϱ�?", vbInformation + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    CommonDialog1.Filter = "Excel 97 - 2003 ���չ��� (*.xls)|*.xls|All Files (*.*)|*.*|Excel Files (*.xls)|*.xls"
    
    'CommonDialog1.Filter = "Excel ���� ���� (*.xlsx)|*.xlsx|All Files (*.*)|*.*|Excel Files (*.xls)|*.xls|Excel 97 - 2003 ���չ��� (*.xls)|*.xls"
    
    CommonDialog1.ShowSave
    
    sFileName = CommonDialog1.Filename
        
    If sFileName = "" Then
        MsgBox "ȭ�ϸ��� �Է��ϼ���!", vbExclamation
        
        Exit Sub
    End If
    
    ClearSpread vasExcel
    
    i = 1
    
    For lRow = 1 To vasList.DataRowCnt
        vasList.Row = lRow
        vasList.col = 1
        
        If vasList.Value = 1 Then
            For iCol = 2 To 50
'                Select Case iCol
'                Case 2      '�Ƿ�����
'                    SetText vasExcel, Trim(GetText(vasList, lRow, iCol)), i, 1
'                Case 3
'                Case Else
'                    SetText vasExcel, Trim(GetText(vasList, lRow, iCol)), i, iCol - 2
'                End Select

                Select Case iCol
                Case 2
                    SetText vasExcel, Trim(GetText(vasList, lRow, iCol)), i, 1
                Case 3
                    SetText vasExcel, Trim(GetText(vasList, lRow, iCol)), i, 2
                Case 4
                    SetText vasExcel, Trim(GetText(vasList, lRow, iCol)), i, 3
                Case 5
                    SetText vasExcel, Trim(GetText(vasList, lRow, iCol)), i, 4
                Case 6
                    SetText vasExcel, Trim(GetText(vasList, lRow, iCol)), i, 5
                Case 7
                    SetText vasExcel, Trim(GetText(vasList, lRow, iCol)), i, 6
                Case 8
                    SetText vasExcel, Trim(GetText(vasList, lRow, iCol)), i, 7
                End Select
            Next iCol
            
            i = i + 1
        End If
    Next lRow
    
    'SaveExcel App.Path & "\" & Format(GetDateShort, "yyyymmdd") & ".xls", vasList
    SaveExcel sFileName, vasExcel
        
    Exit Sub
    
'ErrHandler:
'    '����ڰ� [���] ���߸� �������ϴ�.
'    MsgBox "(" & Err.Number & ") " & Err.Description
'
'    Exit Sub

End Sub

Sub SaveExcel(Filename As String, argSpread As vaSpread)

On Error GoTo ErrHandle

    ' Excel Object Library �� �����մϴ�.
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    
    Dim iRow As Integer
    Dim iCol As Integer


    Set xlApp = CreateObject("Excel.Application")

    Set xlBook = xlApp.Workbooks.Add

    Set xlSheet = xlBook.Worksheets(1)
    
    'xlSheet.Name = Format(GetDateShort, "yyyymmdd")
    
    xlSheet.Columns(1).Select
    xlApp.Selection.NumberFormatLocal = "@"
    
    xlSheet.Columns(2).Select
    xlApp.Selection.NumberFormatLocal = "@"
    
    xlSheet.Columns(3).Select
    xlApp.Selection.NumberFormatLocal = "@"

    xlSheet.Columns(4).Select
    xlApp.Selection.NumberFormatLocal = "@"
    
    xlSheet.Columns(5).Select
    xlApp.Selection.NumberFormatLocal = "@"
    
    xlSheet.Cells(1, 1) = "��ü��ȣ"
    xlSheet.Cells(1, 2) = "�����˻��ڵ�"
    xlSheet.Cells(1, 3) = "ȯ�ڸ�"
    xlSheet.Cells(1, 4) = "�ֹι�ȣ"
    xlSheet.Cells(1, 5) = "����"
    xlSheet.Cells(1, 6) = "����"
    xlSheet.Cells(1, 7) = "íƮ��ȣ"

    For iRow = 1 To argSpread.DataRowCnt
        For iCol = 1 To 10
            argSpread.Row = iRow
            argSpread.col = iCol
            xlSheet.Cells(iRow + 1, iCol) = argSpread.Text
        Next iCol
    Next iRow

    xlBook.SaveAs Filename
    xlApp.Quit
    
    
    MsgBox "���������� ����Ǿ����ϴ�.", vbInformation
    
    Exit Sub
ErrHandle:
    MsgBox "(" & err.Number & ") " & err.Description
    
    Exit Sub

End Sub

Private Sub btnPrint_Click()
    Dim sHead As String
    Dim sHead1 As String
    Dim sFoot As String
    Dim sSlip As String
    Dim sCurDate As String
    Dim sReceNo As String
    Dim sTitle As String
    Dim PageCnt As Integer

    Dim iRow As Integer
    Dim iCol As Integer
    Dim i As Integer
    
On Error GoTo ErrGoto
    
    ClearSpread vasPrint
    
    i = 1
    For iRow = 1 To vasList.DataRowCnt
        vasList.Row = iRow
        vasList.col = 1
        If vasList.Value = 1 Then
            For iCol = 2 To 30
                'SetText vasPrint, Trim(GetText(vasList, iRow, iCol)), i, iCol - 1
                
                Select Case iCol
                Case 5
                    SetText vasPrint, Trim(GetText(vasList, iRow, iCol)), i, 1
                Case 6
                    SetText vasPrint, Trim(GetText(vasList, iRow, iCol)), i, 2
                Case 4
                    SetText vasPrint, Trim(GetText(vasList, iRow, iCol)), i, 3
                Case 7
                    SetText vasPrint, Trim(GetText(vasList, iRow, iCol)), i, 4
                    
                    If Trim(GetText(vasList, iRow, iCol)) <> "" Then
                        SetText vasPrint, Mid(Trim(GetText(vasList, iRow, iCol)), 1, 6), i, 5
                    End If
                Case 9
                    SetText vasPrint, Trim(GetText(vasList, iRow, iCol)), i, 6
                Case 8
                    SetText vasPrint, Trim(GetText(vasList, iRow, iCol)), i, 7
                Case 13
                    SetText vasPrint, Trim(GetText(vasList, iRow, iCol)), i, 8
                Case 14
                    SetText vasPrint, Trim(GetText(vasList, iRow, iCol)), i, 9
                End Select
            Next iCol
            
            i = i + 1
        End If
    Next iRow
    
    If vasPrint.DataRowCnt < 1 Then
        MsgBox "����� �ڷᰡ �����ϴ�.", , "�� ��"
        Exit Sub
    End If

    PageCnt = vasPrint.PrintPageCount
    
    sTitle = "�Ƿ�ȯ�� ����Ʈ"
    
    sCurDate = GetDateFull
       
    sHead = Trim(dtpSDate.Value) & " - " & Trim(dtpEDate.Value)
    
    If optStat(1).Value = True Then
        sHead1 = "������"
    ElseIf optStat(2).Value = True Then
        sHead1 = "����"
    End If
        
        
    vasPrint.PrintOrientation = 2    'SS_PRINTORIENT_LANDSCAPE
    
    
    vasPrint.PrintAbortMsg = "�μ��� �Դϴ� ..."
    vasPrint.PrintJobName = "�Ƿ�ȯ�� ��Ȳ"

    sHead = "/fn""�ü�ü"" /fz""12"" /fb1 /fi0 /fu0 " & "/c" & "�� " & sTitle & " ��" & "/n/n " & _
                "/fn""����ü"" /fz""10"" /fb0 /fi0 /fu0 " & "/c" & "��ȸ���� : " & sHead & "/n/n" & _
                "/fn""����ü"" /fz""10"" /fb0 /fi0 /fu0 " & "/c" & "�˻���� : " & sHead1 & "/rPage /p // " & PageCnt & "/n"
                
    sFoot = "/fn""����ü"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & sCurDate & "/fn""�ü�ü"" /fz""11"" /fb1 /fi0 /fu0 /r" & "�뿵�����ں��� ���ܰ˻����а�"
    vasPrint.PrintHeader = sHead
    vasPrint.PrintFooter = sFoot

    vasPrint.PrintMarginTop = 680
    vasPrint.PrintMarginBottom = 680
'���� SS�� ���Ī���� �����
'    vaslist.PrintMarginLeft = 720
    vasPrint.PrintMarginLeft = 0
    vasPrint.PrintMarginRight = 0
    
    vasPrint.PrintColor = True
    vasPrint.PrintGrid = True

    vasPrint.PrintType = 0  'SS_PRINT_ALL(default)


    vasPrint.PrintShadows = True

    vasPrint.Action = 13 'SS_ACTION_PRINT
    
ErrGoto:
    '����ڰ� ��ҹ�ư�� �������ϴ�.
    Exit Sub
End Sub

Private Sub btnSearch_Click()
    Dim lRow As Long
    
    ClearSpread vasList
    
    SQL = " Select '', ó��ð�, ��������, TO_CHAR(��ü��ȣ), ���Ϲ�ȣ, ����, �������, TO_CHAR(����), �����ڵ�, " & CR & _
          " ���ڵ�, ����, '', ǰ���ڵ�, ǰ���, Ư�����, ó�����ڵ�, '', �ǽ�����, '', '' " & CR & _
          " From �˻��ü2V"
          
'    If cboDate.ListIndex = 1 Then       'ó������
'        SQL = SQL & CR & _
'          "Where �������� Between TO_DATE('" & Replace(Format(CDate(dtpSDate.Value), "mm/dd/yyyy"), "-", "/") & "', 'MM/DD/YYYY') And TO_DATE('" & Replace(Format(CDate(dtpEDate.Value), "mm/dd/yyyy"), "-", "/") & "', 'MM/DD/YYYY') "
'    Else                                '��������
'        SQL = SQL & CR & _
'          "Where �ǽ����� Between TO_DATE('" & Replace(Format(CDate(dtpSDate.Value), "mm/dd/yyyy"), "-", "/") & "', 'MM/DD/YYYY') And TO_DATE('" & Replace(Format(CDate(dtpEDate.Value), "mm/dd/yyyy"), "-", "/") & "', 'MM/DD/YYYY') "
'    End If
    
    If cboDate.ListIndex = 1 Then       'ó������
        SQL = SQL & CR & _
          "Where �������� Between TO_DATE('" & Replace(Format(CDate(dtpSDate.Value), "mm/dd/yyyy"), "-", "/") & "', 'MM/DD/YYYY') And TO_DATE('" & Replace(Format(CDate(dtpEDate.Value), "mm/dd/yyyy"), "-", "/") & "', 'MM/DD/YYYY') "
    Else                                '��������
        SQL = SQL & CR & _
          "Where ó��ð� Between TO_DATE('" & Replace(Format(CDate(dtpSDate.Value), "mm/dd/yyyy"), "-", "/") & "', 'MM/DD/YYYY') And TO_DATE('" & Replace(Format(CDate(dtpEDate.Value), "mm/dd/yyyy"), "-", "/") & "', 'MM/DD/YYYY') "
    End If
    
    SQL = SQL & CR & " And ó�����ڵ� IN ('256','242')"
    
    '�˻����
    If optStat(0).Value = True Then         '���
    
    ElseIf optStat(1).Value = True Then     '����
        SQL = SQL & vbCrLf & _
          "  and ó�������ڵ� = 'I' "
    ElseIf optStat(2).Value = True Then     '�Է�
        SQL = SQL & vbCrLf & _
          "  and ó�������ڵ� = 'R' "
    End If
    
'    SQL = SQL & CR & " Group By ��������, �ǽ�����, ��ü��ȣ, ���Ϲ�ȣ, ����, �������, TO_CHAR(����), �����ڵ�, " & CR & _
'                     " ���ڵ�, ����, ǰ���ڵ�, ǰ���, Ư�����, ó�����ڵ�, �ǽ�����"
    
    SQL = SQL & CR & " Group By ó��ð�, ��������, ��ü��ȣ, ���Ϲ�ȣ, ����, �������, TO_CHAR(����), �����ڵ�, " & CR & _
                     " ���ڵ�, ����, ǰ���ڵ�, ǰ���, Ư�����, ó�����ڵ�, �ǽ�����"
                     
    res = db_select_Vas(SQL, vasList)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    vasList.MaxRows = vasList.DataRowCnt
    
    For lRow = 1 To vasList.DataRowCnt
    
    Next lRow
    
End Sub

Private Sub btnSearch1_Click()
    Dim lRow As Long

    lRow = vasList.DataRowCnt
    
    If lRow > vasList.MaxRows Then
        vasList.MaxRows = lRow + 1
    Else
        lRow = vasList.DataRowCnt + 1
        
        If lRow > vasList.MaxRows Then
            vasList.MaxRows = vasList.DataRowCnt + 1
        End If
    End If
        
    PatInfo txtBarcode, lRow
    
    '��ü��ȣ, �����˻��ڵ�, ȯ�ڸ�, �ֹι�ȣ, ����, ����, íƮ��ȣ
'    SQL = " Select '', ��������, �ǽ�����, TO_CHAR(��ü��ȣ), ���Ϲ�ȣ, ����, �������, TO_CHAR(����), �����ڵ�, " & CR & _
'          " ���ڵ�, ����, '', ǰ���ڵ�, ǰ���, Ư�����, ó�����ڵ�, '', �ǽ�����, '', '' " & CR & _
'          " From �˻��ü2V" & CR & _
'          "Where ��ü��ȣ = '" & Trim(txtBarcode.Text) & "' "
'    SQL = SQL & CR & " And ó�����ڵ� IN ('256','242')"
'
'    If optStat(0).Value = True Then         '���
'
'    ElseIf optStat(1).Value = True Then     '����
'        SQL = SQL & vbCrLf & _
'          "  and ó�������ڵ� = 'I' "
'    ElseIf optStat(2).Value = True Then     '�Է�
'        SQL = SQL & vbCrLf & _
'          "  and ó�������ڵ� = 'R' "
'    End If
'
'    SQL = SQL & CR & " Group By ��������, �ǽ�����, ��ü��ȣ, ���Ϲ�ȣ, ����, �������, TO_CHAR(����), �����ڵ�, " & CR & _
'                     " ���ڵ�, ����, ǰ���ڵ�, ǰ���, Ư�����, ó�����ڵ�, �ǽ�����"
'
'    res = db_select_Vas(SQL, vasList, lRow + 1)
'    If res = -1 Then
'        SaveQuery SQL
'        Exit Sub
'    End If
    
    vasList.MaxRows = vasList.DataRowCnt

End Sub

Private Sub chkAll_Click()
    vasList.Row = -1
    vasList.col = 1
    
    If chkAll.Value = 0 Then
        vasList.Value = 0
    Else
        vasList.Value = 1
    End If
End Sub

Private Sub Form_Load()
    cboDate.ListIndex = 0
    
    dtpSDate.Value = CDate(Date)
    dtpEDate.Value = dtpSDate.Value
    
    ClearSpread vasList
    ClearSpread vasPrint
End Sub

Private Sub txtBarcode_GotFocus()
    SelectFocus txtBarcode
End Sub

Private Sub txtBarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtBarcode <> "" Then
            btnSearch1_Click
            
            txtBarcode = ""
        End If
    End If
End Sub

Private Sub vasList_Click(ByVal col As Long, ByVal Row As Long)

    If Row = 0 Then
        Select Case col
        Case 2      '��������
            vasSort vasList, 2
        Case 3      '�ǽ�����
            vasSort vasList, 3
        Case 4      '��ü��ȣ
            vasSort vasList, 4
        Case 5      'ȯ�ڹ�ȣ
            vasSort vasList, 5
        Case 6      'ȯ�ڼ���
            vasSort vasList, 6
        End Select
    End If
End Sub

Sub PatInfo(asSpecID As String, ByVal asRow As Long)
    Dim sPID As String
    Dim sPName As String
    Dim sSex As String
    Dim sAge As String
    
    Dim sBarCode As String
    
    Dim ResInf As ResInfRec
    Dim Found As Integer
    Dim sResCurKey As String
    Dim sResCmpKey As String
    Dim sRetVal As String
    Dim sCurKey As String
    Dim sResRetVal As String
    
    Dim lsExamCode As String
    Dim lsSpecCode As String
    
    Dim PbsInf As PbsInfRec
    Dim i As Integer
    Dim j As Integer
    Dim ii As Integer
    Dim sAcpDate As String
    Dim sAcpCod As String
    Dim sAcpNum As String
    Dim sOcmNum As String
    Dim sOdrNum As String
    Dim sOdrSeq As String
    
    Dim iFlag As Integer
    
    Dim lRow As Long
    
    Dim iRow As Integer
    
    iRow = asRow
    
    SetText vasList, Trim(asSpecID), asRow, 2
    
    ClearSpread vasOrder

    sBarCode = asSpecID
    
    iFlag = -1
    
If (Len(sBarCode) < 10) Or (Len(sBarCode) = 10 And Left(sBarCode, 1) = "A") Then        '�˻��

    i = 0
    
    sResCmpKey = ""
    
    If Len(sBarCode) < 10 Then      '��������|�˻���Ʈ|������ȣ
        sAcpNum = sBarCode
        sAcpCod = "REF"
        sAcpDate = "20" & Mid(sBarCode, 2, 6)
        
        sResCurKey = sAcpDate & Chr(5) & sAcpCod & Chr(5) & SetSpace(sAcpNum, 10) & Chr(5)
    Else
        ClearSpread vasTemp1
        
        SQL = " Select barcode "
        SQL = SQL & CR & " From barcodeinfo "
        SQL = SQL & CR & "where barcode1 = '" & sBarCode & "' "
        SQL = SQL & CR & " Group By barcode"
        res = db_select_Vas(SQL, vasTemp1)
        
        If vasTemp1.DataRowCnt > 0 Then
            iFlag = 1
        Else
        
            SQL = "Select barcode, ocmnum, odrnum, odrseq, acpdte, acpcod, acpnum, spmcode "
            SQL = SQL & " from barcodeinfo "
            SQL = SQL & "where barcode = '" & sBarCode & "' "
            res = db_select_Col(SQL)
            If Trim(gReadBuf(0)) = sBarCode Then
                sAcpDate = Trim(gReadBuf(4))
                If Len(sAcpDate) = 10 Then
                    sAcpDate = Format(sAcpDate, "YYYYMMDD")
                End If
                
                sAcpCod = Trim(gReadBuf(5))
                sAcpNum = Trim(gReadBuf(6))
                
                sResCurKey = sAcpDate & Chr(5) & sAcpCod & Chr(5) & SetSpace(sAcpNum, 10) & Chr(5)
                'Save_Raw_Data "[���ڵ�� ������������]" & sResCurKey
            Else
                Exit Sub
            End If
        End If
    End If
    
    If iFlag = 1 Then
        For lRow = 1 To vasTemp1.DataRowCnt
            sAcpDate = ""
            sAcpCod = ""
            sAcpNum = ""
            
            SQL = "Select barcode, ocmnum, odrnum, odrseq, acpdte, acpcod, acpnum, spmcode "
            SQL = SQL & " from barcodeinfo "
            SQL = SQL & "where barcode = '" & Trim(GetText(vasTemp1, lRow, 1)) & "' "
            res = db_select_Col(SQL)
            If Trim(gReadBuf(0)) <> "" Then
                sAcpDate = Trim(gReadBuf(4))
                If Len(sAcpDate) = 10 Then
                    sAcpDate = Format(sAcpDate, "YYYYMMDD")
                End If
                
                sAcpCod = Trim(gReadBuf(5))
                sAcpNum = Trim(gReadBuf(6))
                
                sResCurKey = sAcpDate & Chr(5) & sAcpCod & Chr(5) & SetSpace(sAcpNum, 10) & Chr(5)
                
                sResCurKey = mSetNext("ResInf", sResCurKey)
                Do
                    sResCurKey = mReadNext("ResInf", sResCurKey, sResCmpKey, sResRetVal)
                    'Debug.Print sResRetVal
                    'Save_Raw_Data sResRetVal
                    
                    If sResCurKey = "" Then Exit Do
                    
                    If piece(sResRetVal, Chr(5), 3) <> SetSpace(sAcpNum, 10) Then Exit Do
                    
                    lsSpecCode = piece(sResRetVal, Chr(5), 4)
                    lsExamCode = piece(sResRetVal, Chr(5), 5)
                    
                    'vaslist.SetText colSpecimen, asRow, lsSpecCode
                    
                    If Trim(piece(sResRetVal, Chr(5), 16)) <> "" Then       '������翩��
                        lsExamCode = ""
                        lsSpecCode = ""
                    Else
                        i = i + 1
                        If vasOrder.MaxRows < i Then
                            vasOrder.MaxRows = i
                        End If
            
                        vasOrder.SetText 1, i, lsSpecCode
                        vasOrder.SetText 2, i, lsExamCode
                    End If
                    
                    'íƮ��ȣ
                    sPID = piece(sResRetVal, Chr(5), 7)
                    
                    If Trim(GetText(vasList, asRow, 8)) = "" Then
                        vasList.SetText 8, asRow, piece(sResRetVal, Chr(5), 7)
                    
                        sCurKey = piece(sResRetVal, Chr(5), 7) & Chr(5)
                        sCurKey = mSetReadEqual("PbsInf", sCurKey, sRetVal)
                        If sCurKey <> "" Then
                            vasList.SetText 4, asRow, Trim(piece(sRetVal, Chr(5), 2))   'ȯ�ڸ�
                            vasList.SetText 5, asRow, Trim(piece(sRetVal, Chr(5), 3))   '�ֹι�ȣ
                            
                            CalAgeSex Trim(piece(sRetVal, Chr(5), 3)), CDate(Date)
                            
                            vasList.SetText 6, asRow, gPatGen.Sex
                            vasList.SetText 7, asRow, gPatGen.Age
                            'vasList.SetText colReceNo, asRow, Trim(piece(sResRetVal, Chr(5), 3))
                        End If
                    End If
                Loop
                
                ii = 1
                For j = 1 To vasOrder.DataRowCnt
                    If Trim(GetText(vasOrder, j, 2)) <> "" Then
                        
                        If ii = 1 Then
                            SetText vasList, Trim(GetText(vasOrder, j, 2)), iRow, 3
                            ii = ii + 1
                        Else
                            iRow = iRow + 1
                            If vasList.MaxRows < iRow Then
                                vasList.MaxRows = iRow
                            End If
                            
                            SetText vasList, Trim(GetText(vasList, asRow, 2)), iRow, 2
                            SetText vasList, Trim(GetText(vasOrder, j, 2)), iRow, 3
                            SetText vasList, Trim(GetText(vasList, asRow, 4)), iRow, 4
                            SetText vasList, Trim(GetText(vasList, asRow, 5)), iRow, 5
                            SetText vasList, Trim(GetText(vasList, asRow, 6)), iRow, 6
                            SetText vasList, Trim(GetText(vasList, asRow, 7)), iRow, 7
                            SetText vasList, Trim(GetText(vasList, asRow, 8)), iRow, 8
                        End If
                    End If
                Next j
            Else
                Exit Sub
            End If
        Next lRow
    Else
        sResCurKey = mSetNext("ResInf", sResCurKey)
        Do
            sResCurKey = mReadNext("ResInf", sResCurKey, sResCmpKey, sResRetVal)
            'Debug.Print sResRetVal
            Save_Raw_Data sResRetVal
            
            If sResCurKey = "" Then Exit Do
            
            If piece(sResRetVal, Chr(5), 3) <> SetSpace(sAcpNum, 10) Then Exit Do
            
            lsSpecCode = piece(sResRetVal, Chr(5), 4)
            lsExamCode = piece(sResRetVal, Chr(5), 5)
            
            'vaslist.SetText colSpecimen, asRow, lsSpecCode
            
            If Trim(piece(sResRetVal, Chr(5), 16)) <> "" Then       '������翩��
                lsExamCode = ""
                lsSpecCode = ""
            Else
                i = i + 1
                If vasOrder.MaxRows < i Then
                    vasOrder.MaxRows = i
                End If

                vasOrder.SetText 1, i, lsSpecCode
                vasOrder.SetText 2, i, lsExamCode
            End If
            
            'íƮ��ȣ
            sPID = piece(sResRetVal, Chr(5), 7)
            
            If Trim(GetText(vasList, asRow, 8)) = "" Then
                vasList.SetText 8, asRow, piece(sResRetVal, Chr(5), 7)
            
                sCurKey = piece(sResRetVal, Chr(5), 7) & Chr(5)
                sCurKey = mSetReadEqual("PbsInf", sCurKey, sRetVal)
                If sCurKey <> "" Then
                    vasList.SetText 4, asRow, Trim(piece(sRetVal, Chr(5), 2))   'ȯ�ڸ�
                    vasList.SetText 5, asRow, Trim(piece(sRetVal, Chr(5), 3))   '�ֹι�ȣ
                    
                    CalAgeSex Trim(piece(sRetVal, Chr(5), 3)), CDate(Date)
                    
                    vasList.SetText 6, asRow, gPatGen.Sex
                    vasList.SetText 7, asRow, gPatGen.Age
                    'vasList.SetText colReceNo, asRow, Trim(piece(sResRetVal, Chr(5), 3))
                    
'                    SetText vasOrder, "test1", 1, 2
'                    SetText vasOrder, "test2", 2, 2
                End If
            End If
        Loop
        
        ii = 1
        For j = 1 To vasOrder.DataRowCnt
            If Trim(GetText(vasOrder, j, 2)) <> "" Then
                
                If ii = 1 Then
                    SetText vasList, Trim(GetText(vasOrder, j, 2)), iRow, 3
                    ii = ii + 1
                Else
                    iRow = iRow + 1
                    If vasList.MaxRows < iRow Then
                        vasList.MaxRows = iRow
                    End If
                    
                    SetText vasList, Trim(GetText(vasList, asRow, 2)), iRow, 2
                    SetText vasList, Trim(GetText(vasOrder, j, 2)), iRow, 3
                    SetText vasList, Trim(GetText(vasList, asRow, 4)), iRow, 4
                    SetText vasList, Trim(GetText(vasList, asRow, 5)), iRow, 5
                    SetText vasList, Trim(GetText(vasList, asRow, 6)), iRow, 6
                    SetText vasList, Trim(GetText(vasList, asRow, 7)), iRow, 7
                    SetText vasList, Trim(GetText(vasList, asRow, 8)), iRow, 8
                End If
            End If
        Next j
    End If
    
ElseIf Len(sBarCode) = 10 Or Left(sBarCode, 1) <> "A" Then      '����
    SQL = "Select barcode, ocmnum, odrnum, odrseq, acpdte, acpcod, acpnum, spmcode "
    SQL = SQL & " from barcodeinfo "
    SQL = SQL & "where barcode = '" & sBarCode & "' "
    res = db_select_Col(SQL)
    If Trim(gReadBuf(0)) = sBarCode Then
        sOcmNum = Trim(gReadBuf(1))
        sOdrNum = Trim(gReadBuf(2))
        sOdrSeq = Trim(gReadBuf(3))

        sResCurKey = SetSpace(sOcmNum, 10) & Chr(5) & SetSpace(sOdrNum, 4) & Chr(5) & SetSpace(sOdrSeq, 5) & Chr(5)
    Else
        Exit Sub
    End If

    i = 0

    sResCmpKey = ""

    'sResCurKey = Format(CDate(Trim(DTPicker1.Value)), "yyyymmdd") & Chr(5) & gEquipSlip & Chr(5) & SetSpace(sBarCode, 10) & Chr(5)

    sResCurKey = mSetNext("ResInfOcmOdrOdr", sResCurKey)
    Do
        sResCurKey = mReadNext("ResInfOcmOdrOdr", sResCurKey, sResCmpKey, sResRetVal)
        Debug.Print sResRetVal
        Save_Raw_Data sResRetVal

        If sResCurKey = "" Then Exit Do

        If piece(sResRetVal, Chr(5), 6) <> SetSpace(sOcmNum, 10) Then Exit Do
        If piece(sResRetVal, Chr(5), 39) <> SetSpace(sOdrNum, 4) Then Exit Do

        lsSpecCode = piece(sResRetVal, Chr(5), 4)
        lsExamCode = piece(sResRetVal, Chr(5), 5)

        'vaslist.SetText colSpecimen, asRow, lsSpecCode

        If Trim(piece(sResRetVal, Chr(5), 16)) <> "" Then
            lsExamCode = ""
            lsSpecCode = ""
        Else
            i = i + 1
            If vasOrder.MaxRows < i Then
                vasOrder.MaxRows = i
            End If

            vasOrder.SetText 1, i, lsSpecCode
            vasOrder.SetText 2, i, lsExamCode
        End If

        'íƮ��ȣ
        sPID = piece(sResRetVal, Chr(5), 7)

        If Trim(GetText(vasList, asRow, 8)) = "" Then
            vasList.SetText 8, asRow, piece(sResRetVal, Chr(5), 7)          'íƮ��ȣ

            sCurKey = piece(sResRetVal, Chr(5), 7) & Chr(5)
            sCurKey = mSetReadEqual("PbsInf", sCurKey, sRetVal)
            If sCurKey <> "" Then
                vasList.SetText 4, asRow, Trim(piece(sRetVal, Chr(5), 2))   'ȯ�ڸ�
                vasList.SetText 5, asRow, Trim(piece(sRetVal, Chr(5), 3))   '�ֹι�ȣ
                
                CalAgeSex Trim(piece(sRetVal, Chr(5), 3)), CDate(Date)

                vasList.SetText 6, asRow, gPatGen.Sex
                vasList.SetText 7, asRow, gPatGen.Age
                'vasList.SetText colReceNo, asRow, Trim(piece(sResRetVal, Chr(5), 3))
            End If
        End If
    Loop

    ii = 1
    For j = 1 To vasOrder.DataRowCnt
        If Trim(GetText(vasOrder, j, 2)) <> "" Then
            
            If ii = 1 Then
                SetText vasList, Trim(GetText(vasOrder, j, 2)), iRow, 3
                ii = ii + 1
            Else
                iRow = iRow + 1
                If vasList.MaxRows < iRow Then
                    vasList.MaxRows = iRow
                End If
                
                SetText vasList, Trim(GetText(vasList, asRow, 2)), iRow, 2
                SetText vasList, Trim(GetText(vasOrder, j, 2)), iRow, 3
                SetText vasList, Trim(GetText(vasList, asRow, 4)), iRow, 4
                SetText vasList, Trim(GetText(vasList, asRow, 5)), iRow, 5
                SetText vasList, Trim(GetText(vasList, asRow, 6)), iRow, 6
                SetText vasList, Trim(GetText(vasList, asRow, 7)), iRow, 7
                SetText vasList, Trim(GetText(vasList, asRow, 8)), iRow, 8
            End If
        End If
    Next j
                
Else

End If

End Sub

