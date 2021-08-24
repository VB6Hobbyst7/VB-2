VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmTmpPtB 
   BackColor       =   &H00E8EEEE&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "수탁검사"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox txtPtid 
      Height          =   375
      Left            =   3900
      TabIndex        =   0
      Top             =   195
      Width           =   2085
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00EBF3ED&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   4545
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   2850
      Width           =   1320
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00FEF5F3&
      Caption         =   "확인(&O)"
      Height          =   510
      Left            =   3210
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   2850
      Width           =   1320
   End
   Begin VB.TextBox txtNm 
      Height          =   375
      Left            =   3555
      TabIndex        =   1
      Top             =   930
      Width           =   2430
   End
   Begin VB.TextBox txtSSn2 
      Height          =   375
      Left            =   4950
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1770
      Width           =   1035
   End
   Begin VB.TextBox txtSSN1 
      Height          =   375
      Left            =   3570
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1770
      Width           =   975
   End
   Begin MSComctlLib.ListView lvwPtList 
      Height          =   3285
      Left            =   60
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   180
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   5794
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "환자ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "환자명"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "주민번호"
         Object.Width           =   2540
      EndProperty
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   360
      Index           =   11
      Left            =   2445
      TabIndex        =   9
      Top             =   195
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
      Left            =   2445
      TabIndex        =   10
      Top             =   1770
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
      Height          =   360
      Index           =   1
      Left            =   2445
      TabIndex        =   11
      Top             =   930
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
      Caption         =   "성명"
      Appearance      =   0
   End
   Begin VB.Label lblPtid 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   1  '단일 고정
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   3555
      TabIndex        =   8
      Top             =   195
      Width           =   330
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E8EEEE&
      Caption         =   "~"
      Height          =   255
      Left            =   4665
      TabIndex        =   6
      Top             =   1860
      Width           =   135
   End
End
Attribute VB_Name = "frmTmpPtB"
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
    
    txtPtid.Text = Format(txtPtid.Text, "00000000")
    
    blnSave = TmpPtSave("A" & txtPtid.Text, Trim(txtNm), txtSSN1 & txtSSn2)
    
    
    If blnSave = True Then
        strPara = "A" & txtPtid & COL_DIV & txtNm & COL_DIV & SDA_String(txtSSN1 & txtSSn2)
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
    Dim SSQL As String
    Dim RS   As Recordset
    Dim iTmx As ListItem
    Dim ii   As Long
    
    SSQL = " SELECT " & F_PTID & " as ptid," & F_PTNM & " as ptnm," & F_SSN2 & " as ssn " & _
           " FROM " & T_HIS001 & " a " & _
           " WHERE  a." & F_PTID & " like 'A%' AND " & F_PTNM & "='테스트' ORDER BY " & F_PTID
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        With lvwPtList
            Do Until RS.EOF
            
                If ii > 20 Then Exit Sub
                
                Set iTmx = .ListItems.Add(, , "" & RS.Fields("ptid").Value)
                
                
                iTmx.SubItems(1) = "" & RS.Fields("ptnm").Value
                
                If RS.Fields("SSN").Value & "" > 6 Then
                    iTmx.SubItems(2) = Mid("" & RS.Fields("SSN").Value, 1, 6) & "-" & _
                    Mid("" & RS.Fields("ssn").Value, 7)
                
                End If
                RS.MoveNext
                ii = ii + 1
                
                
            Loop
        End With
    End If
    Set RS = Nothing
End Sub

Private Function SSNChek(ByVal ssn As String) As Boolean
    
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



Private Sub lvwPtList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lblPtid.Caption = Mid(Item.Text, 1, 1)
    txtPtid.Text = Mid(Item.Text, 2)
    txtNm.Text = Item.SubItems(1)
    If Len(Item.SubItems(2)) > 6 Then
        txtSSN1 = Mid(Item.SubItems(2), 1, 6)
        txtSSn2 = Mid(Replace(Item.SubItems(2), "-", ""), 7)
    Else
        txtSSN1 = Item.SubItems(2)
        txtSSn2 = ""
    End If
End Sub

Private Sub txtNm_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txtPtId_KeyPress(KeyAscii As Integer)
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
    
    'SSQL = " update " & T_HIS001 & " set " & DBW(F_PTNM, ptnm, 3) & _
           DBW("resno1", ssn1, 2) & DBW("resno2", ssn2, 2) & _
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





