VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmSupplemental 
   BackColor       =   &H00DDE1DE&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Supplemental Report"
   ClientHeight    =   4875
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10065
   Icon            =   "frmSupplemental.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPreview 
      Caption         =   "미리보기(&R)"
      Height          =   375
      Left            =   1110
      TabIndex        =   7
      Top             =   570
      Width           =   1230
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "출력(&P)"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   570
      Width           =   945
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "종료(&X)"
      Height          =   375
      Left            =   8955
      TabIndex        =   1
      Top             =   570
      Width           =   960
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "저장(&S)"
      Height          =   375
      Left            =   7935
      TabIndex        =   0
      Top             =   570
      Width           =   960
   End
   Begin RichTextLib.RichTextBox txtSupp 
      Height          =   3735
      Left            =   105
      TabIndex        =   3
      Top             =   1020
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   6588
      _Version        =   393217
      BackColor       =   16054772
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmSupplemental.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Crystal.CrystalReport crReport 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "종합검증/판독 보고서"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lblMfyDt 
      BackStyle       =   0  '투명
      Height          =   165
      Left            =   4710
      TabIndex        =   5
      Top             =   660
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "수정일 : "
      Height          =   180
      Left            =   3945
      TabIndex        =   4
      Top             =   660
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Supplemental Report"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E48372&
      Height          =   180
      Left            =   3705
      TabIndex        =   2
      Top             =   285
      Width           =   2085
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00DBF2FD&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   390
      Left            =   3135
      Shape           =   4  '둥근 사각형
      Top             =   195
      Width           =   3285
   End
End
Attribute VB_Name = "frmSupplemental"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_PtId As String
Private m_BedinDt As String
Private m_RptDt As String
Private m_SaveFg As Boolean

Private SuppExist As Boolean

Public Property Get ptid() As String
    ptid = m_PtId
End Property
Public Property Let ptid(ByVal vData As String)
    m_PtId = vData
End Property

Public Property Get RptDt() As String
    RptDt = m_RptDt
End Property
Public Property Let RptDt(ByVal vData As String)
    m_RptDt = vData
End Property

Public Property Get BedinDt() As String
    BedinDt = m_BedinDt
End Property
Public Property Let BedinDt(ByVal vData As String)
    m_BedinDt = vData
End Property

Private Sub cmdExit_Click()
    Unload Me
    Set frmSupplemental = Nothing
End Sub

Private Sub cmdPreview_Click()
    Call PrtReport(1)
End Sub

Private Sub cmdPrint_Click()
    Call PrtReport(0)
End Sub

Private Sub cmdSave_Click()
    
    Dim SqlStmt As String
    
    If m_PtId = "" Then
        MsgBox "저장할 대상이 없습니다.", vbInformation
        Exit Sub
    End If
    
    If SuppExist Then
        SqlStmt = " update " & T_LAB504 & " set " & _
                  " mfydt  = '" & Format(Now, CS_DateDbFormat) & "', " & _
                  " txtrst = '" & Trim(txtSupp.Text) & "' " & _
                  " where " & DBW("ptid  = ", m_PtId) & _
                  " and   " & DBW("rptdt = ", m_RptDt)
    Else
        SqlStmt = " insert into " & T_LAB504 & " (ptid, rptdt, mfyseq, mfydt, txtrst) values (" & _
                  DBV("ptid", m_PtId) & ", " & DBV("rptdt", m_RptDt) & ", 1, '" & Format(Now, CS_DateDbFormat) & "', '" & _
                  Trim(txtSupp.Text) & "')"
        SuppExist = True
    End If
    
    On Error GoTo Err_Trap
    
    DBConn.BeginTrans
    DBConn.Execute SqlStmt
    DBConn.CommitTrans
    
    m_SaveFg = True
    If Trim(txtSupp.Text) <> "" Then
        cmdPrint.Enabled = True
    Else
        cmdPrint.Enabled = False
    End If
    
    MsgBox "정상적으로 저장되었습니다.", vbInformation
    
    Exit Sub
    
Err_Trap:
'    Call Error_Routine
    DBConn.RollbackTrans
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub GetSuppText()

    Dim SqlStmt As String
    Dim Rs As Recordset
    
    SqlStmt = " select * from " & T_LAB504 & _
              " where " & DBW("ptid=", m_PtId) & _
              " and   " & DBW("rptdt=", m_RptDt) & _
              " order by mfyseq"
'    Set Rs = OpenRecordSet(SqlStmt)
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Not Rs.EOF Then
        lblMfyDt.Caption = Format("" & Rs.Fields("MfyDt").Value, CS_DateLongMask)
        txtSupp.Text = Trim("" & Rs.Fields("TxtRst").Value)
        SuppExist = True
        cmdPrint.Enabled = True
    Else
        SuppExist = False
        cmdPrint.Enabled = False
    End If
    
'    Rs.RsClose
    Set Rs = Nothing
    
    m_SaveFg = True
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim Resp As VbMsgBoxResult
    
    If Not m_SaveFg Then
        Resp = MsgBox("변경된 데이타를 저장하지 않고 종료하시겠습니까?", vbQuestion + vbYesNo, "메세지")
        If Resp = vbNo Then
            Cancel = True
        End If
    End If
    
End Sub

Private Sub txtSupp_Change()

    m_SaveFg = False

End Sub

Public Sub PrtReport(ByVal pOption As Integer)

    Dim hwndPreviewWindow As Long
    Dim SqlStmt As String

    Me.MousePointer = 11

On Error GoTo PRINT_ERROR
    
    SqlStmt = "SELECT  S2HIS001." & F_PTNM & ", S2HIS001." & F_ADDRESS & ", " & _
              "        S2LAB502.ptid, S2LAB501.bedindt, S2LAB501.wardid, S2LAB501.hosilid, S2LAB501.deptcd, S2LAB501.rptdt, " & _
              "        S2LAB502.age, S2LAB502.agediv, S2LAB502.sex, " & _
              "        S2LAB505.doctnm, S2LAB505.doctno, S2LAB505.certno, " & _
              "        S2HIS006." & F_IENM & ", " & _
              "        S2LAB504.txtrst "
    SqlStmt = SqlStmt & " From " & _
                        T_LAB505 & " S2LAB505, " & _
                        T_LAB501 & " S2LAB501, " & _
                        T_HIS001 & " S2HIS001, " & _
                        T_HIS006 & " S2HIS006, " & _
                        T_LAB502 & " S2LAB502, " & _
                        T_LAB504 & " S2LAB504"
    SqlStmt = SqlStmt & " WHERE " & _
                                  DBW("S2LAB501.ptid = ", m_PtId) & _
                        " AND " & DBW("S2LAB501.bedindt = ", m_BedinDt) & _
                        " AND S2LAB505.doctid = S2LAB501.rptid  " & _
                        " AND S2LAB501.ptid = S2HIS001." & F_PTID & _
                        " AND " & DBJ("S2HIS006." & F_ICD & " =* S2LAB501.disease") & _
                        " AND S2LAB501.ptid = S2LAB502.ptid " & _
                        " AND S2LAB501.rptdt = S2LAB502.rptdt " & _
                        " AND S2LAB502.ptid = S2LAB504.ptid " & _
                        " AND S2LAB502.rptdt = S2LAB504.rptdt "
    With crReport
        
        GetConnInfo
        .Connect = "DSN=" & medGetP(GetConnInfo, 1, ";") & ";UID=" & medGetP(GetConnInfo, 2, ";") & ";PWD=" & medGetP(GetConnInfo, 3, ";") & ";"
        .ReportFileName = InstallDir & "Lis\Rpt\LabSupp.rpt" 'LISRptPath & "\LabSupp.rpt"
        If pOption = 1 Then
            .Destination = crptToWindow  '0 ' Window
            .WindowLeft = 0
            .WindowTop = 0
            .WindowState = crptMaximized
        Else
            .Destination = crptToPrinter
        End If
        .ParameterFields(1) = "bedindt;" & Format(m_BedinDt, CS_DateLongMask) & ";TRUE"
        .SQLQuery = SqlStmt
        .Action = 1 ' Print
    End With
        
PRINT_ERROR:
    Me.MousePointer = 0

End Sub
