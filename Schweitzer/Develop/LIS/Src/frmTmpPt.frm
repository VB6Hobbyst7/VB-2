VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmTmpPt 
   BackColor       =   &H00E8EEEE&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "종합검진 환자등록"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   Icon            =   "frmTmpPt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox txtid 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   120
      Width           =   1395
   End
   Begin VB.TextBox txtSSN1 
      Height          =   375
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtSSn2 
      Height          =   375
      Left            =   2580
      MaxLength       =   7
      TabIndex        =   2
      Top             =   600
      Width           =   1035
   End
   Begin VB.TextBox txtNm 
      Height          =   405
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00FEF5F3&
      Caption         =   "확인(&O)"
      Height          =   510
      Left            =   2520
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   1035
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00EBF3ED&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   3840
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   1035
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   360
      Index           =   11
      Left            =   75
      TabIndex        =   7
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   635
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "등록번호"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   360
      Index           =   0
      Left            =   75
      TabIndex        =   8
      Top             =   585
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   635
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "주민번호"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   375
      Index           =   1
      Left            =   2610
      TabIndex        =   9
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "성명"
      Appearance      =   0
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E8EEEE&
      Caption         =   "~"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   720
      Width           =   135
   End
End
Attribute VB_Name = "frmTmpPt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mvarRTop  As Long
Private mvarRLeft As Long
Private mvarPtid  As String
Private mvarPtNm  As String
Private mvarSSN  As String

Public Event OKButtonClick(ByVal strPtInfo As String)

Public Property Let RTop(ByVal vData As Long)
    mvarRTop = vData
End Property
Public Property Let Rleft(ByVal vData As Long)
    mvarRLeft = vData
End Property
Public Property Let PtId(ByVal vData As String)
    mvarPtid = vData
End Property
Public Property Let ptnm(ByVal vData As String)
    mvarPtNm = vData
End Property
Public Property Let ssn(ByVal ssn As String)
    mvarSSN = ssn
End Property

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    'Dim objSql  As New clsPtInformation
    
    Dim blnSave As Boolean
    Dim strPara As String
    Dim SSQL As String
    
    
    '이름/생년월일/성,나이
    If SSNChek(txtSSN1 & txtSSn2) = False Then Exit Sub
    
    
    blnSave = TmpPtSave(UCase(Trim(txtid)), Trim(txtNm), txtSSN1 & txtSSn2)
    
    If blnSave = True Then
        strPara = txtNm & COL_DIV & SDA_String(txtSSN1 & txtSSn2)
        RaiseEvent OKButtonClick(strPara)
    Else
        MsgBox " 건진환자 등록 오류입니다.", vbInformation + vbOKOnly, "건진 ID등록"
    End If
    Unload Me
    
    
    
End Sub


Private Sub Form_Activate()
    frmTmpPt.Top = mvarRTop
    frmTmpPt.Left = mvarRLeft
End Sub

Private Sub Form_Load()
    txtid = mvarPtid
    txtNm = mvarPtNm
    If Len(mvarSSN) >= 6 Then
        txtSSN1 = Mid(mvarSSN, 1, 6)
        txtSSn2 = Mid(mvarSSN, 7)
    Else
        txtSSN1 = "": txtSSn2 = ""
    End If
End Sub

Private Function SSNChek(ByVal ssn As String) As Boolean
    '주민등록번호 체크
'    Dim intN1 As Integer, intN2 As Integer, intN3 As Integer, intN4 As Integer, intN5 As Integer
'    Dim intN6 As Integer, intN7 As Integer, intN8 As Integer, intN9 As Integer, intN10 As Integer
'    Dim intN11 As Integer, intN12 As Integer, intN13 As Integer, intSsn As Integer, intSsn1 As Integer, intSsn2 As Integer
'
'    intN1 = Val(Mid(ssn, 1, 1))
'    intN2 = Val(Mid(ssn, 2, 1))
'    intN3 = Val(Mid(ssn, 3, 1))
'    intN4 = Val(Mid(ssn, 4, 1))
'    intN5 = Val(Mid(ssn, 5, 1))
'    intN6 = Val(Mid(ssn, 6, 1))
'    intN7 = Val(Mid(ssn, 7, 1))
'    intN8 = Val(Mid(ssn, 8, 1))
'    intN9 = Val(Mid(ssn, 9, 1))
'    intN10 = Val(Mid(ssn, 10, 1))
'    intN11 = Val(Mid(ssn, 11, 1))
'    intN12 = Val(Mid(ssn, 12, 1))
'    intN13 = Val(Mid(ssn, 13, 1))
'
'    intSsn = (intN1 * 2) + (intN2 * 3) + (intN3 * 4) + (intN4 * 5) + (intN5 * 6) + (intN6 * 7) + (intN7 * 8) _
'        + (intN8 * 9) + (intN9 * 2) + (intN10 * 3) + (intN11 * 4) + (intN12 * 5)
'    intSsn1 = 11 - (intSsn Mod 11)
'
'    If intSsn1 > 9 Then
'        intSsn2 = intSsn1 Mod 10
'        If intSsn2 = intN13 Then
'            SSNChek = True
'            Exit Function
'        Else
'            MsgBox "주민번호가 잘못 되었습니다.다시 확인 하십시요.", vbCritical, "등록실패"
'            SSNChek = False
'            Exit Function
'        End If
'    Else
'        If intSsn1 = intN13 Then
'            SSNChek = True
'            Exit Function
'        Else
'            MsgBox "주민번호가 잘못 되었습니다.다시 확인 하십시요.", vbCritical, "등록실패"
'            SSNChek = False
'            Exit Function
'        End If
'    End If
    
    If Len(ssn) < 7 Then
        MsgBox "주민번호가 잘못 되었습니다.다시 확인 하십시요.", vbCritical, "등록실패"
        SSNChek = False
        Exit Function
    End If
    If Mid(ssn, 7, 1) = 0 Or Mid(ssn, 7, 1) > 4 Then
        MsgBox "주민번호가 잘못 되었습니다.다시 확인 하십시요.", vbCritical, "등록실패"
        SSNChek = False
        Exit Function
    End If
    If Mid(ssn, 3, 2) = "00" Or Mid(ssn, 3, 2) > 12 Then
        MsgBox "주민번호가 잘못 되었습니다.다시 확인 하십시요.", vbCritical, "등록실패"
        SSNChek = False
        Exit Function
    End If
    If Mid(ssn, 5, 2) = "00" Or Mid(ssn, 5, 2) > 31 Then
        MsgBox "주민번호가 잘못 되었습니다.다시 확인 하십시요.", vbCritical, "등록실패"
        SSNChek = False
        Exit Function
    End If
    
    SSNChek = True
End Function

Private Sub txtNm_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txtSSN1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txtSSn2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub


 Private Function SDA_String(ByVal ssn As String) As String
    Dim strTmp As String
    Dim strSEX As String
    Dim strAge As String
    Dim strDOB As String
    
    Dim strYY  As String
    Dim strMM  As String
    Dim strDD  As String
    
    strYY = Mid(ssn, 1, 2)
    strMM = Mid(ssn, 3, 2)
    strDD = Mid(ssn, 5, 2)
    
    If Val(strMM) < 1 Then strMM = "01"
    If Val(strMM) > 12 Then strMM = "12"
    If Val(strDD) < 1 Then strDD = "01"
    If Val(strDD) > 31 Then strDD = "31"
    
    If IsDate(strYY & "-" & strMM & "-" & strDD) = False Then
        strDD = "01"
    End If
    
    strSEX = "기타": strAge = "": strDOB = ""
    
    If ssn <> "" Then
        strTmp = Mid(ssn, 7, 1)
        Select Case strTmp
            Case "0": strSEX = "여": strDOB = "18" & strYY & "-" & strMM & "-" & strDD
            Case "1": strSEX = "남": strDOB = "19" & strYY & "-" & strMM & "-" & strDD
            Case "2": strSEX = "여": strDOB = "19" & strYY & "-" & strMM & "-" & strDD
            Case "3": strSEX = "남": strDOB = "20" & strYY & "-" & strMM & "-" & strDD
            Case "4": strSEX = "여": strDOB = "20" & strYY & "-" & strMM & "-" & strDD
            Case Else: strSEX = "남": strDOB = "19" & strYY & "-" & strMM & "-" & strDD
        End Select
        
        If Len(ssn) > 6 Then
            If strDOB <> "" Then
                strAge = medFindAge(Replace(strDOB, "-", ""), "Y")
            End If
        Else
            strAge = ""
        End If
        SDA_String = strSEX & COL_DIV & strDOB & COL_DIV & strAge
    Else
        SDA_String = "" & COL_DIV & "" & COL_DIV & ""
    End If
End Function

Private Function TmpPtSave(ByVal PtId As String, ByVal ptnm As String, ByVal ssn As String) As Boolean
    Dim SSQL As String
    Dim ssn1 As String
    Dim ssn2 As String
    
    ssn1 = Mid(ssn, 1, 6)
    ssn2 = Mid(ssn, 7)
    
    On Error GoTo SAVE_ERROR
    DBConn.BeginTrans
    'SSQL = " update " & T_HIS001 & " set " & DBW(F_PTNM, ptnm, 3) & DBW("resno1=", ssn1, 1) & DBW("resno2=", ssn2) & _
          " WHERE " & DBW(F_PTID, PtId, 2)
          
    '20131214 DB 암호화
    SSQL = "   update appatbat set patname = '" & ptnm & "' " & vbCRLF
    SSQL = SSQL & " , resno1 = CryptIT.encrypt('" & ssn1 & "','pmc1898') " & vbCRLF
    SSQL = SSQL & " , resno2 = CryptIT.encrypt('" & ssn2 & "','pmc1898') " & vbCRLF
    SSQL = SSQL & " WHERE PTNO = '" & PtId & "' "
    
    DBConn.Execute SSQL
    DBConn.CommitTrans
    TmpPtSave = True

    Exit Function
SAVE_ERROR:
    DBConn.RollbackTrans
End Function



