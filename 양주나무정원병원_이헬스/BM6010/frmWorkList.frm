VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmWorkList 
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "��ũ����Ʈ ��ȸ"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17205
   Icon            =   "frmWorkList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   17205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin VB.TextBox txtBarcode 
      Alignment       =   2  '��� ����
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   180
      Width           =   1755
   End
   Begin VB.TextBox txtRack 
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
      Left            =   8280
      TabIndex        =   16
      Text            =   "1"
      Top             =   150
      Width           =   525
   End
   Begin VB.TextBox txtNum 
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
      Left            =   8820
      TabIndex        =   15
      Text            =   "A"
      Top             =   150
      Width           =   525
   End
   Begin FPSpread.vaSpread vasPrint 
      Height          =   3315
      Left            =   3120
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   11580
      _Version        =   393216
      _ExtentX        =   20426
      _ExtentY        =   5847
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
      MaxCols         =   9
      SpreadDesigner  =   "frmWorkList.frx":000C
   End
   Begin VB.CommandButton cmdWorkPrint 
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12300
      TabIndex        =   13
      Top             =   180
      Width           =   1395
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "Check1"
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   900
      TabIndex        =   11
      Top             =   810
      Width           =   225
   End
   Begin VB.TextBox txtSeq 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   15960
      TabIndex        =   9
      Text            =   "1"
      Top             =   210
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13740
      TabIndex        =   8
      Top             =   180
      Width           =   1395
   End
   Begin VB.CommandButton cmdDownClose 
      Caption         =   "Down >> Close"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10860
      TabIndex        =   7
      Top             =   180
      Width           =   1395
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "��ȸ"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4380
      TabIndex        =   2
      Top             =   180
      Width           =   1035
   End
   Begin VB.CommandButton cmdDownLoad 
      Caption         =   "Down"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9420
      TabIndex        =   1
      Top             =   180
      Width           =   1395
   End
   Begin MSComCtl2.DTPicker dtpStartDt 
      Height          =   315
      Left            =   1290
      TabIndex        =   3
      Top             =   180
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   129171457
      CurrentDate     =   40457
   End
   Begin MSComCtl2.DTPicker dtpStopDt 
      Height          =   315
      Left            =   2850
      TabIndex        =   4
      Top             =   180
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   129171457
      CurrentDate     =   40457
   End
   Begin FPSpread.vaSpread vasWorkList 
      Height          =   5865
      Left            =   300
      TabIndex        =   12
      Top             =   720
      Width           =   16695
      _Version        =   393216
      _ExtentX        =   29448
      _ExtentY        =   10345
      _StockProps     =   64
      ButtonDrawMode  =   4
      ColHeaderDisplay=   0
      ColsFrozen      =   16
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   17
      MaxRows         =   20
      MoveActiveOnFocus=   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmWorkList.frx":1AA6
   End
   Begin VB.Label Label3 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "R/P"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   7800
      TabIndex        =   17
      Top             =   270
      Width           =   375
   End
   Begin VB.Label Label2 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "Seq"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   15450
      TabIndex        =   10
      Top             =   270
      Width           =   375
   End
   Begin VB.Label Label1 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "��ȸ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   330
      TabIndex        =   6
      Top             =   270
      Width           =   780
   End
   Begin VB.Label Label7 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2730
      TabIndex        =   5
      Top             =   270
      Width           =   105
   End
End
Attribute VB_Name = "frmWorkList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAll_Click()
    Dim iRow As Long
    
    If chkAll.Value = 1 Then
        For iRow = 1 To vasWorkList.DataRowCnt
            vasWorkList.Row = iRow
            vasWorkList.Col = 1
            
            vasWorkList.Value = 1
        Next iRow
    ElseIf chkAll.Value = 0 Then
        For iRow = 1 To vasWorkList.DataRowCnt
            vasWorkList.Row = iRow
            vasWorkList.Col = 1
            
            vasWorkList.Value = 0
        Next iRow
    End If
    
End Sub

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub

Private Sub cmdDownClose_Click()
    
    Call cmdDownLoad_Click
    
    Call cmdClose_Click
    
End Sub

Private Sub cmdDownLoad_Click()
    Dim intVasRow As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    Dim j  As Integer
    Dim strOrd  As String
    Dim intOrdCnt As Integer
    
    j = 0
    With vasWorkList
        For intRow = 1 To .MaxRows
            .Row = intRow
            .Col = colCheckBox
            If .Value = 1 Then
                frmInterface.vasID.MaxRows = frmInterface.vasID.MaxRows + 1
                DoEvents
                intVasRow = frmInterface.vasID.MaxRows
                
                If GetText(vasWorkList, intRow, colBARCODE) = "" Then
                    frmInterface.vasID.MaxRows = frmInterface.vasID.MaxRows - 1
                    Exit Sub
                End If
                
                strOrd = GetText(vasWorkList, intRow, colDISKNO)
                intOrdCnt = Len(strOrd) / 4
                
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colSpecNo), intVasRow, colSpecNo)
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colCheckBox), intVasRow, colCheckBox)
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colHOSPDATE), intVasRow, colHOSPDATE)
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colBARCODE), intVasRow, colBARCODE)
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colCHARTNO), intVasRow, colCHARTNO)
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colPID), intVasRow, colPID)
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colINOUT), intVasRow, colINOUT)
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colPNAME), intVasRow, colPNAME)
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colPSEX), intVasRow, colPSEX)
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colPAGE), intVasRow, colPAGE)
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colPOSNO), intVasRow, colPOSNO)
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colDISKNO), intVasRow, colDISKNO)
                Call SetText(frmInterface.vasID, intOrdCnt, intVasRow, colOCNT)
                
                For intCol = colState + 1 To .MaxCols
                    .Row = intVasRow
                    .Col = intCol
                    If .BackColor = vbYellow Then
                        Call SetBackColor(frmInterface.vasID, intVasRow, intVasRow, intCol, intCol, 255, 255, 0)
                    End If
                Next
                
                
                
                .Col = 1
                .Value = "0"
            End If
        Next
        frmInterface.vasID.RowHeight(-1) = 12
    End With

    frmInterface.dtpStartDt.Value = dtpStartDt.Value
    frmInterface.dtpStopDt.Value = dtpStopDt.Value
    
    frmInterface.txtRack.Text = txtRack.Text
    frmInterface.txtNum.Text = txtNum.Text



'    Dim i As Integer
'
'    If KeyAscii = vbKeyReturn Then
'        For i = 1 To vasWorkList.MaxRows
'            vasWorkList.Row = i
'            vasWorkList.Col = 1
'            If vasWorkList.Value = "1" Then
'                If Trim(txtPos.Text) = "" Then
'                    txtPos.Text = "1"
'                End If
'                Call SetText(frmInterface.vasworklist, Format(txtPos.Text, "0000"), i, 0)
'                txtPos.Text = Format(txtPos.Text + 1, "0000")
'            End If
'        Next
'    End If
End Sub

Private Sub cmdSearch_Click()

    Call GetWorkList_EHEALTH(Format(dtpStartDt, "yyyymmdd"), Format(dtpStopDt, "yyyymmdd"))
    
End Sub

Private Sub GetWorkList_MCC(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim strBarcode    As String
    Dim blnSame     As Boolean
    
    If pBarNo = "" Then
        vasWorkList.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    vasWorkList.ReDraw = False
    

'''          SQL = "SELECT DISTINCT ORD_YMD, BCODE_NO, RECEPT_NO, PTNT_NO,PTNT_NM,AGE,SEX,ORD_CD" & vbCr
'''    SQL = SQL & "  FROM MCCSI.H7LIS_BCODE_ORD " & vbCr
'''    SQL = SQL & " WHERE ORD_YMD between '" & pFrDt & "' AND '" & pToDt & "'" & vbCr
'''    SQL = SQL & "   AND ORD_CD IN (" & gAllExam & ") " & vbCr
'''    SQL = SQL & "   AND RESULT_TYPE = '20'" & vbLf & vbCr
'''    SQL = SQL & "  ORDER BY ORD_YMD,RECEPT_NO,BCODE_NO "

          SQL = "SELECT DISTINCT  a.ptnt_no, c.ptnt_nm, a.recept_no, a.spc_cd, " & vbCr
    SQL = SQL & " (select codeval1 from pm_mst_div_key1 where codediv = 'LAB01' and codekey1 = a.spc_cd) as spc_nm " & vbCr
    SQL = SQL & "      , a.sts_cd, a.acc_ymd, a.ord_cd " & vbCr
    SQL = SQL & "  FROM h3lab_result a, h1opdin b, hz_mst_ptnt c " & vbCr
    SQL = SQL & " WHERE a.recept_no = b.recept_no " & vbCr
    SQL = SQL & "   AND a.sutak_cd = ''" & vbCr
    SQL = SQL & "   AND a.ptnt_no  = c.ptnt_no" & vbCr
    SQL = SQL & "   AND a.sts_cd   = 'A'" & vbCr                                                               ' A:���� R:���"
    SQL = SQL & "   AND a.acc_ymd between '" & pFrDt & "' AND '" & pToDt & "'" & vbCr
    SQL = SQL & "   AND a.ord_cd IN (" & gAllExam & ") " & vbCr
    SQL = SQL & " Order by recept_no " & vbCr

    Call SetSQLData("��ũ��ȸ", SQL)

    '-- Record Count ������
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        
        Do Until RS.EOF
            iCnt = iCnt + 1
            With vasWorkList
                .ReDraw = False
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasWorkList, i, colHOSPDATE)
                    strBarcode = GetText(vasWorkList, i, colBARCODE)
                    If Trim(RS("acc_ymd")) = strDate And Trim(RS("recept_no")) = strBarcode Then
                        blnSame = True
                    End If
                    
                    For intCol = colState + 1 To vasWorkList.MaxCols
                        If Trim(RS.Fields("ord_cd")) = gArrEquip(intCol - colState, 3) Then
                            vasWorkList.Row = .MaxRows
                            vasWorkList.Col = intCol
                            vasWorkList.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                Next
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    
                    SetText vasWorkList, txtSeq.Text, .MaxRows, colSpecNo
                    SetText vasWorkList, "1", .MaxRows, colCheckBox
                    SetText vasWorkList, Trim(RS.Fields("acc_ymd")) & "", .MaxRows, colHOSPDATE
                    SetText vasWorkList, Trim(RS.Fields("recept_no")) & "", .MaxRows, colBARCODE
                    'SetText vasWorkList, Trim(RS.Fields("ptnt_no")) & "", .MaxRows, colCHARTNO
                    SetText vasWorkList, Trim(RS.Fields("ptnt_no")) & "", .MaxRows, colPID
                    SetText vasWorkList, Trim(RS.Fields("ptnt_nm")) & "", .MaxRows, colPNAME
                    'SetText vasWorkList, Trim(RS.Fields("AGE")) & "", .MaxRows, colPAGE
                    'SetText vasWorkList, Trim(RS.Fields("SEX")) & "", .MaxRows, colPSEX
                    
                    txtSeq.Text = txtSeq.Text + 1
                    
                    For intCol = colState + 1 To vasWorkList.MaxCols
                        If Trim(RS.Fields("ord_cd")) = gArrEquip(intCol - colState, 3) Then
                            vasWorkList.Row = .MaxRows
                            vasWorkList.Col = intCol
                            vasWorkList.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                End If
                blnSame = False
            End With
                        
            RS.MoveNext
        Loop
        chkAll.Value = "1"
    Else
        'StatusBar1.Panels(3).Text = "��ȸ ����ڰ� �����ϴ�."
        chkAll.Value = "0"
    End If
    
    RS.Close
    
    '-- ���α׷����� �ݱ�
    Unload frmProgress
    
    vasWorkList.RowHeight(-1) = 12
    vasWorkList.ReDraw = True
    
End Sub

Private Sub GetWorkList_EHEALTH(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim strBarcode    As String
    Dim blnSame     As Boolean
    Dim strTests    As String
    Dim intOrdCnt   As Integer
    
'    If pBarNo = "" Then
'        vasWorkList.MaxRows = 0
'        intRow = 0
'    End If
    
    blnSame = False
    vasWorkList.ReDraw = False
    
          SQL = "SELECT DISTINCT b.OBODORDT"    '�Է���
    SQL = SQL & ", a.APATMRNO"                  '��Ϲ�ȣ
    SQL = SQL & ", b.OBODCASE"                  '������ȣ
    SQL = SQL & ", b.OBODORNO"                  'ORDER NUMBER
    SQL = SQL & ", b.OBODORSQ"                  'ORDER SEQUENCE
    SQL = SQL & ", b.OBODIOGB"                  '��/�� I=�Կ�/O=�ܷ�
    SQL = SQL & ", a.APATNAME"                  'ȯ�ڼ���
    SQL = SQL & ", a.APATPSEX"                  '����(M/F)
    SQL = SQL & ", a.APATJMN1"                  '�ֹι�ȣ(�����)
    SQL = SQL & ", b.OBODCODE"                  '�����ڵ�
    SQL = SQL & ", c.OBSUCODE"                  '�˻��ڵ�
    SQL = SQL & ", c.OBSUSUBC" & vbCr           '�˻��ڵ�SUB
    
    SQL = SQL & "  FROM ABPATMST a"             'ȯ�ڱ⺻����
    SQL = SQL & "      ,OBODRMTM b"            'ó�泻�� Table
    SQL = SQL & "      ,OBSURSTM c " & vbCr     '�˻���(��ġ���) Table
    
    SQL = SQL & " WHERE a.APATMRNO = b.OBODMRNO " & vbCr                                '��Ϲ�ȣ,����ȣ
    SQL = SQL & "   AND a.APATMRNO = c.OBSUMRNO " & vbCr                                '��Ϲ�ȣ,����ȣ
    SQL = SQL & "   AND b.OBODCASE = c.OBSUCASE " & vbCr                                '������ȣ
    SQL = SQL & "   AND b.OBODORNO = c.OBSUORNO " & vbCr                                'ORDER NUMBER
    SQL = SQL & "   AND b.OBODORSQ = c.OBSUORSQ " & vbCr                                'ORDER SEQUENCE
    SQL = SQL & "   AND (c.OBSURSLT IS NULL OR c.OBSURSLT = '')" & vbCr                 '�˻���
    SQL = SQL & "   AND RTRIM(LTRIM(c.OBSUCODE)) + '|' + RTRIM(LTRIM(c.OBSUSUBC)) IN (" & gAllExam & ") " & vbCr    '�˻��ڵ� + '|' + OBSUSUBC
    SQL = SQL & "   AND b.OBODSTAT = 'AC' " & vbCr                                      '�ʼ� �⺻ = 'OE', ä���� = 'AC'
    
    If pBarNo <> "" Then
        SQL = SQL & "   AND a.APATMRNO = '" & pBarNo & "'"
    Else
        SQL = SQL & "   AND b.OBODORDT between '" & pFrDt & "' AND '" & pToDt & "'" & vbCr  '�Է���
    End If
    
    SQL = SQL & " Order by b.OBODORDT,a.APATMRNO,b.OBODORNO,b.OBODORSQ " & vbCr

    Call SetSQLData("��ũ��ȸ", SQL)

    '-- Record Count ������
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        intOrdCnt = 0
        Do Until RS.EOF
            iCnt = iCnt + 1
            With vasWorkList
                .ReDraw = False
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasWorkList, i, colHOSPDATE)
                    strBarcode = GetText(vasWorkList, i, colBARCODE)
                    If Trim(RS("OBODORDT")) = strDate And Trim(RS("APATMRNO")) = strBarcode Then
                        blnSame = True
                    End If
                    
                    If blnSame = True Then
                        For intCol = colState + 1 To vasWorkList.MaxCols
                            If Trim(RS.Fields("OBSUCODE")) & "|" & Trim(RS.Fields("OBSUSUBC")) = gArrEquip(intCol - colState, 3) Then
                                vasWorkList.Row = .MaxRows
                                vasWorkList.Col = intCol
                                vasWorkList.BackColor = vbYellow
                                vasWorkList.Text = "��"
                                Exit For
                            End If
                        Next
                    End If
                Next
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    
                    txtNum = txtNum + 1
                    If txtNum = "85" Then
                        txtRack = txtRack + 1
                        txtNum = 1
                    End If
                    
                    SetText vasWorkList, txtSeq.Text, .MaxRows, colSpecNo
                    SetText vasWorkList, "1", .MaxRows, colCheckBox
                    SetText vasWorkList, Trim(RS.Fields("OBODORDT")) & "", .MaxRows, colHOSPDATE
                    SetText vasWorkList, Trim(RS.Fields("APATMRNO")) & "", .MaxRows, colBARCODE
                    SetText vasWorkList, Trim(RS.Fields("OBODCASE")) & "", .MaxRows, colCHARTNO
                    SetText vasWorkList, Trim(RS.Fields("OBODORNO")) & "", .MaxRows, colPID
                    SetText vasWorkList, IIf(Trim(RS.Fields("OBODIOGB")) & "" = "I", "�Կ�", "�ܷ�"), .MaxRows, colINOUT
                    SetText vasWorkList, Trim(RS.Fields("APATNAME")) & "", .MaxRows, colPNAME
                    SetText vasWorkList, Trim(RS.Fields("APATPSEX")) & "", .MaxRows, colPSEX
                    SetText vasWorkList, Trim(RS.Fields("APATJMN1")) & "", .MaxRows, colPAGE
                    
                    SetText vasWorkList, Format(txtRack, "00") & "-" & Format(txtNum, "00"), .MaxRows, colPOSNO
                    
                    strTests = GetOrderExamCode(gEquip, .MaxRows)

                    SetText vasWorkList, strTests, .MaxRows, colDISKNO
                    'txtSeq.Text = txtSeq.Text + 1
                    
                    
                    For intCol = colState + 1 To vasWorkList.MaxCols
                        If Trim(RS.Fields("OBSUCODE")) & "|" & Trim(RS.Fields("OBSUSUBC")) = gArrEquip(intCol - colState, 3) Then
                            vasWorkList.Row = .MaxRows
                            vasWorkList.Col = intCol
                            vasWorkList.BackColor = vbYellow
                            vasWorkList.Text = "��"
                            Exit For
                        End If
                    Next
                    
                End If
                
                
                blnSame = False
            End With
                        
            RS.MoveNext
        Loop
        chkAll.Value = "1"
    Else
        'StatusBar1.Panels(3).Text = "��ȸ ����ڰ� �����ϴ�."
        chkAll.Value = "0"
    End If
    
    RS.Close
    
    '-- ���α׷����� �ݱ�
    'Unload frmProgress
    
    vasWorkList.RowHeight(-1) = 12
    vasWorkList.ReDraw = True
    
End Sub

Private Sub cmdWorkPrint_Click()
Dim iRow As Integer
Dim iCol As Integer
Dim j As Integer

Dim sCurDate As String
Dim sSerDate As String
Dim sHead As String
Dim sFoot As String
Dim strTestNm As String

    ClearSpread vasPrint

    j = 1

    vasPrint.RowHeight(-1) = 20
    
    For iRow = 1 To vasWorkList.DataRowCnt
        vasWorkList.Row = iRow
        vasWorkList.Col = colCheckBox

        If vasWorkList.Value = 1 Then
            SetText vasPrint, Trim(GetText(vasWorkList, iRow, colBARCODE)), j, 1     '��ü��ȣ
            SetText vasPrint, Trim(GetText(vasWorkList, iRow, colCHARTNO)), j, 2     'ȯ�ڹ�ȣ
            SetText vasPrint, Trim(GetText(vasWorkList, iRow, colPNAME)), j, 3     'ȯ���̸�

            SetText vasPrint, Trim(GetText(vasWorkList, iRow, colPSEX)), j, 4     '����
            SetText vasPrint, Trim(GetText(vasWorkList, iRow, colPAGE)), j, 5     '����
            
            strTestNm = ""
            For iCol = colState + 1 To vasWorkList.MaxCols
                vasWorkList.Row = iRow
                vasWorkList.Col = iCol
                If vasWorkList.Text = "��" Then
                    strTestNm = strTestNm & Trim(GetText(vasWorkList, 0, iCol)) & "/"
                End If
            Next
            SetText vasPrint, strTestNm, j, 6     '�˻��׸�

            j = j + 1
        End If
    Next iRow

    If vasPrint.DataRowCnt < 1 Then
        MsgBox "����� �ڷᰡ �����ϴ�.", , "�� ��"
        Exit Sub
    End If

    sCurDate = GetDateFull

    sSerDate = Trim(dtpStartDt.Value) & " - " & Trim(dtpStopDt.Value)

    vasPrint.PrintOrientation = 1   ' SS_PRINTORIENT_PORTRAIT
    vasPrint.PrintAbortMsg = "�μ��� �Դϴ� ..."
    vasPrint.PrintJobName = "WorkList ���"


    sHead = "/fn""�ü�ü"" /fz""12"" /fb1 /fi0 /fu0 " & "/c" & "�� ��ȭ�� WorkList ��" & "/n/n " & _
            "/fn""����ü"" /fz""10"" /fb0 /fi0 /fu0 " & "/c" & "ó������ : " & dtpStartDt & " ~ " & dtpStopDt

    sFoot = "/fn""����ü"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & sCurDate & "/fn""�ü�ü"" /fz""11"" /fb1 /fi0 /fu0 /r" & " ������������ �˻��"

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

'Set printing range
    vasPrint.PrintType = 0  'SS_PRINT_ALL(default)

    vasPrint.PrintShadows = True

    vasPrint.Action = 13 'SS_ACTION_PRINT

'    vasWorkList.PrintOrientation = PrintOrientationLandscape '�������
'    vasWorkList.Action = ActionPrint
    
End Sub

Private Sub Form_Load()
    
    dtpStartDt.Value = frmInterface.dtpStartDt.Value
    dtpStopDt.Value = frmInterface.dtpStopDt.Value
'    txtSeq.Text = frmInterface.txtSeq.Text
    
    vasWorkList.MaxRows = 0
    
    txtRack = "1"
    txtNum = 0
    
End Sub







Private Sub txtBarcode_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    Dim intRow As Long
    
    If KeyAscii = 13 Then
        If Not IsNumeric(txtBarcode) Then
            MsgBox "���ڵ��ȣ�� ���ڸ� �Է��� �����մϴ�."
            txtBarcode = ""
            Exit Sub
        End If
        
        If Len(txtBarcode) <> 8 Then
            MsgBox "���ڵ� �ڸ����� Ȯ���ϼ���"
            txtBarcode = ""
            Exit Sub
        End If
        
        If Trim(txtBarcode) <> "" Then
            Call GetWorkList_EHEALTH(Format(dtpStartDt, "yyyymmdd"), Format(dtpStopDt, "yyyymmdd"), Trim(txtBarcode.Text))
        End If
        
'        vasID.RowHeight(-1) = 12
        txtBarcode.Text = ""
    End If
End Sub

Private Sub txtSeq_KeyPress(KeyAscii As Integer)
    Dim intRow As Integer
    
    If KeyAscii = vbKeyReturn Then
        
        For intRow = vasWorkList.ActiveRow To vasWorkList.MaxRows
            Call SetText(vasWorkList, Val(txtSeq.Text), intRow, colSpecNo)
            txtSeq.Text = txtSeq.Text + 1
        Next
        
        txtSeq.Text = Format(txtSeq.Text, "0000")
    
    End If


End Sub

Private Sub vasWorkList_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim pGrid_Point As Integer
    Dim sBarcode As String
    Dim sChartNo As String
    Dim intCol   As Integer
    Dim strOrd  As String
    Dim intOrdCnt As Integer
    
    If Row = 0 Then Exit Sub
    
    With vasWorkList
        pGrid_Point = SeqSearch(frmInterface.vasID, GetText(vasWorkList, Row, colBARCODE), colBARCODE)

        If pGrid_Point = 0 Then
            pGrid_Point = SeqNullSearch(frmInterface.vasID, Trim(.Text), colBARCODE)
            If pGrid_Point = 0 Then
                frmInterface.vasID.MaxRows = frmInterface.vasID.MaxRows + 1
                pGrid_Point = frmInterface.vasID.MaxRows
            End If
            .RowHeight(-1) = 12
        End If

        strOrd = GetText(vasWorkList, Row, colDISKNO)
        intOrdCnt = Len(strOrd) / 4

        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colSpecNo), pGrid_Point, colSpecNo)
        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colCheckBox), pGrid_Point, colCheckBox)
        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colHOSPDATE), pGrid_Point, colHOSPDATE)
        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colBARCODE), pGrid_Point, colBARCODE)
        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colPID), pGrid_Point, colPID)
        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colCHARTNO), pGrid_Point, colCHARTNO)
        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colPNAME), pGrid_Point, colPNAME)
        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colPSEX), pGrid_Point, colPSEX)
        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colPAGE), pGrid_Point, colPAGE)

        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colINOUT), pGrid_Point, colINOUT)
        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colPOSNO), pGrid_Point, colPOSNO)
        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colDISKNO), pGrid_Point, colDISKNO)
        Call SetText(frmInterface.vasID, intOrdCnt, pGrid_Point, colOCNT)


        frmInterface.vasID.RowHeight(-1) = 12

        For intCol = colState + 1 To .MaxCols
            .Row = Row
            .Col = intCol
            
            If .BackColor = vbYellow Then
            'If .BackColor = vbYellow Then
            'If .Text = "��" Then
                Call SetBackColor(frmInterface.vasID, pGrid_Point, pGrid_Point, intCol, intCol, 255, 255, 0)
            End If
        Next


    End With
End Sub
