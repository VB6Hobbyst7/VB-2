VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frm423RICnt 
   BackColor       =   &H00DBE6E6&
   ClientHeight    =   9780
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11445
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frm423.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9780
   ScaleWidth      =   11445
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00EBF3ED&
      Caption         =   "�� ��(&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '�׷���
      TabIndex        =   8
      Tag             =   "0"
      Top             =   8505
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00EBF3ED&
      Caption         =   "�� ��(&S)"
      Height          =   510
      Left            =   8175
      Style           =   1  '�׷���
      TabIndex        =   7
      Tag             =   "0"
      Top             =   8505
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   75
      TabIndex        =   2
      Top             =   45
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   556
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
      Caption         =   "�����н� �˻�Ǽ� ��ȸ"
      LeftGab         =   100
   End
   Begin FPSpread.vaSpread tblRiCnt 
      Height          =   6915
      Left            =   75
      TabIndex        =   4
      Tag             =   "10114"
      Top             =   1530
      Width           =   10770
      _Version        =   196608
      _ExtentX        =   18997
      _ExtentY        =   12197
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColor       =   14737632
      MaxCols         =   6
      MaxRows         =   21
      OperationMode   =   2
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "frm423.frx":000C
      StartingColNumber=   2
      VirtualRows     =   24
      VisibleCols     =   5
      VisibleRows     =   21
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   945
      Left            =   75
      TabIndex        =   3
      Top             =   285
      Width           =   10785
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00EBF3ED&
         Caption         =   "�� ȸ(&Q)"
         Height          =   510
         Left            =   3045
         Style           =   1  '�׷���
         TabIndex        =   1
         Tag             =   "0"
         Top             =   255
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   360
         Left            =   1290
         TabIndex        =   0
         Top             =   300
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         _Version        =   393216
         Format          =   75169793
         CurrentDate     =   37509
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   2
         Left            =   105
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   285
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
   Begin MedControls1.LisLabel lblPro 
      Height          =   315
      Left            =   75
      TabIndex        =   5
      Top             =   1200
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   556
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
      Caption         =   "��ȸ���"
      LeftGab         =   100
   End
End
Attribute VB_Name = "frm423RICnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*--------------------------------------------------------
'*  Programed by : �̻��
'*  Last Update : 2002-09-12
'*--------------------------------------------------------
Option Explicit

Public Event FormClose()

Private Sub cmdExit_Click()
    Unload Me
    RaiseEvent FormClose
End Sub

Private Sub cmdQuery_Click()
    Dim strDate As String           '������ ��¥
    
    strDate = Format(dtpdate.Value, "yyyyMMdd")
    Call GetRiCnt(strDate)
End Sub

Private Sub cmdSave_Click()
    Dim objPro As jProgressBar.clsProgress
    Dim strDate As String
    Dim strDate1 As String
    Dim strWorkArea As String
    Dim strTestCd As String
    Dim strBussDiv As String
    Dim strCount As Long
    Dim intTemp As Integer
    Dim i As Long
    
    If tblRiCnt.DataRowCnt = 0 Then Exit Sub
    
    strDate = Format(dtpdate.Value, "yyyyMMdd")
    strDate1 = Format(dtpdate.Value, "yyyy�� MM�� dd��")
    If ExistData(strDate) = True Then
        intTemp = MsgBox(strDate1 & "�� �̹� �����Ǿ����ϴ�. �縶�� �Ͻðڽ��ϱ�?", _
                                  vbYesNo + vbInformation, "����")
        If intTemp = vbNo Then Exit Sub
        
        '�̹� ����� ������ ����
        Call DeleteRiCnt(strDate)
    End If
    
    'ProgressBar ó��
    Set objPro = Nothing
    Set objPro = New jProgressBar.clsProgress
    
    With objPro
        .Container = Me
        .Left = lblPro.Left
        .Top = lblPro.Top
        .Width = lblPro.Width
        .Height = lblPro.Height
        .Message = "�������Դϴ�..."
        .Max = tblRiCnt.DataRowCnt
        
'        .Choice = True
'        .SetMyForm Me
'        .XPos = lblPro.Left
'        .YPos = lblPro.Top
'        .XWidth = lblPro.Width
'        .YHeight = lblPro.Height
'        .Appearance = aPlate
'        .Msg = "�������Դϴ�..."
'        .Max = tblRiCnt.DataRowCnt
    End With
    
    '������ ����
    With tblRiCnt
        For i = 1 To .DataRowCnt - 1
            .Row = i
            .Col = 1: strDate = Format(.Value, "yyyyMMdd")
            .Col = 2: strWorkArea = .Value
            .Col = 4: strTestCd = .Value
            .Col = 5
            Select Case .Value
                Case "�ܷ�": strBussDiv = "1"
                Case "����": strBussDiv = "2"
            End Select
            .Col = 6: strCount = Val(.Value)
            
            objPro.Value = i
            Call InsertRiCnt(strDate, strWorkArea, strTestCd, strBussDiv, strCount)
        Next i
    End With
    
    'Spread Clear
    Call medClearTable(tblRiCnt)
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub Form_Load()
    dtpdate.Value = GetSystemDate
End Sub

'*--------------------------------------------------------
'*  ��� : �����н� ������, WA, �˻��ڵ庰 �˻�Ǽ� ��ȸ, �������忡 ��� ���
'*  Parameter
'*      qDate : ������
'*--------------------------------------------------------
Private Sub GetRiCnt(ByVal qDate As String)
    Dim Rs As Recordset
    Dim objPro As jProgressBar.clsProgress
    Dim SQL As String
    Dim strDate As String       '������
    Dim strWorkArea As String   'Workarea
    Dim strBussDiv As String    '�ܷ�/����
    Dim lngTotal As Long
    Dim i As Long

    SQL = "SELECT A.VFYDT, A.WORKAREA, C.TESTNM, C.TESTCD, B.BUSSDIV, COUNT(*) AS COUNT FROM "
    SQL = SQL & T_LAB001 & " C, "
    SQL = SQL & T_LAB101 & " B, "
    SQL = SQL & T_LAB302 & " A "
    SQL = SQL & "WHERE (A.WORKAREA='OR' OR A.WORKAREA='RI') AND "
    SQL = SQL & DBW("A.VFYDT=", qDate) & " AND "
    SQL = SQL & "A.PTID=B.PTID AND A.ORDDT=B.ORDDT AND A.ORDNO=B.ORDNO AND "
    SQL = SQL & "A.TESTCD=C.TESTCD "
    SQL = SQL & "GROUP BY A.VFYDT, A.WORKAREA, C.TESTNM, C.TESTCD, B.BUSSDIV"
    
On Error GoTo Error
    Set Rs = New Recordset
    Rs.Open SQL, dbconn
    
    'ProgressBar ó��
    Set objPro = Nothing
    Set objPro = New jProgressBar.clsProgress
    With objPro
        .Container = Me
        .Left = lblPro.Left
        .Top = lblPro.Top
        .Width = lblPro.Width
        .Height = lblPro.Height
        .Max = Rs.RecordCount
        .Message = "�˻����Դϴ�..."
'        .Choice = True
'        .SetMyForm Me
'        .XPos = lblPro.Left
'        .YPos = lblPro.Top
'        .XWidth = lblPro.Width
'        .YHeight = lblPro.Height
'        .Appearance = aPlate
'        .Msg = "�˻����Դϴ�..."
'        .Max = Rs.RecordCount
    End With
    
    '��ȸ����� Spread�� ���
    Call medClearTable(tblRiCnt)
    If Not Rs.EOF Then
        With tblRiCnt
            Do Until Rs.EOF
                If .DataRowCnt >= .MaxRows Then
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                Else
                    .Row = .DataRowCnt + 1
                End If
                
                .Col = 1    '������
                strDate = Format(Rs.Fields("vfydt").Value & "", "####-##-##")
                .Value = strDate
                .Col = 2: .Value = Rs.Fields("workarea").Value & ""     'WorkArea
                .Col = 3: .Value = Rs.Fields("testnm").Value & ""       '�˻��
                .Col = 4: .Value = Rs.Fields("testcd").Value & ""       '�˻��ڵ�
                .Col = 5    '�ܷ�/����
                Select Case Rs.Fields("bussdiv").Value & ""
                    Case "1":
                        strBussDiv = "�ܷ�"
                    Case "2":
                        strBussDiv = "����"
                End Select
                .Value = strBussDiv
                .Col = 6: .Value = Rs.Fields("count").Value & ""        '��ȸ�Ǽ�
                lngTotal = lngTotal + .Value
                
                i = i + 1
                objPro.Value = i
                Rs.MoveNext
            Loop
            
            .Row = .DataRowCnt + 1: .Col = 5
            .Value = "�� ��"
            .Col = 6: .Value = lngTotal
        End With
    End If

    '�޸� ����
    Set objPro = Nothing
    Set Rs = Nothing
    Exit Sub
    
Error:
'    If rs.DBerror = True Then
'        MsgBox dbconn.Errors.Item(1).Description
'    End If
    MsgBox Err.Description, vbExclamation
    
    Set Rs = Nothing
    Set objPro = Nothing
End Sub

'*--------------------------------------------------------
'*  ��� : �ش���¥�� �̹� ������ �����Ͱ� �ִ��� Ȯ��
'*  Parameter
'*      qDate : ������
'*  Return
'*      True : �̹� �����Ͱ� ���� ���
'*      False : �����Ͱ� ���°��
'*--------------------------------------------------------
Private Function ExistData(ByVal qDate As String) As Boolean
    Dim Rs As Recordset
    Dim SQL As String

On Error GoTo Error
    SQL = "SELECT vfydt FROM S2RICNT WHERE " & DBW("vfydt=", qDate)
    Set Rs = New Recordset
    Rs.Open SQL, dbconn
    
    If Rs.EOF Then
        ExistData = False
    Else
        ExistData = True
    End If
    Exit Function

Error:
    MsgBox Err.Description, vbExclamation
End Function

'*--------------------------------------------------------
'*  ��� : �ش��� ��ȸ����� S2RICNT ���̺� ����
'*  Parameter
'*      qDate : ������
'*--------------------------------------------------------
Private Sub InsertRiCnt(ByVal qDate As String, ByVal qWA As String, ByVal qTestCd As String, _
                                ByVal qBussDiv As String, ByVal qCount As Long)
    Dim SQL As String
    
On Error GoTo Error
    SQL = "INSERT INTO S2RICNT VALUES ('"
    SQL = SQL & qDate & "', '"
    SQL = SQL & qWA & "', '"
    SQL = SQL & qTestCd & "', '"
    SQL = SQL & qBussDiv & "', "
    SQL = SQL & qCount & ")"
    
    dbconn.Execute SQL
    Exit Sub
    
Error:
    MsgBox Err.Description, vbExclamation
End Sub

'*--------------------------------------------------------
'*  ��� : �ش��� ��ȸ����� S2RICNT ���̺��� ����
'*  Parameter
'*      qDate : ������
'*--------------------------------------------------------
Private Sub DeleteRiCnt(ByVal qDate As String)
    Dim SQL As String

On Error GoTo Error
    SQL = "DELETE FROM S2RICNT WHERE" & DBW("vfydt=", qDate)
    dbconn.Execute SQL
    Exit Sub

Error:
    MsgBox Err.Description, vbExclamation
End Sub
