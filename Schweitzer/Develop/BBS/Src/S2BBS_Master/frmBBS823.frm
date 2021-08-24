VERSION 5.00
Begin VB.Form frmBBS823 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "접수번호 생성기준"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   420
      Left            =   3840
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   4020
      Width           =   1260
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5130
      Style           =   1  '그래픽
      TabIndex        =   0
      Tag             =   "128"
      Top             =   4020
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   975
      Left            =   2160
      TabIndex        =   5
      Top             =   2940
      Width           =   5895
      Begin VB.CheckBox chkAccNo 
         BackColor       =   &H00DBE6E6&
         Caption         =   "연별 생성"
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   3
         Top             =   420
         Value           =   1  '확인
         Width           =   1095
      End
      Begin VB.CheckBox chkAccNo 
         BackColor       =   &H00DBE6E6&
         Caption         =   "월별 생성"
         Height          =   255
         Index           =   1
         Left            =   2340
         TabIndex        =   2
         Top             =   420
         Width           =   1095
      End
      Begin VB.CheckBox chkAccNo 
         BackColor       =   &H00DBE6E6&
         Caption         =   "일자별 생성"
         Height          =   255
         Index           =   0
         Left            =   780
         TabIndex        =   1
         Top             =   420
         Width           =   1335
      End
   End
   Begin VB.Label lblAccNo 
      Caption         =   "Label1"
      Height          =   255
      Left            =   4860
      TabIndex        =   6
      Top             =   2580
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmBBS823"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private onPgm As Boolean

Private Sub chkAccNo_Click(Index As Integer)
    Dim i As Long
    
    If onPgm = True Then Exit Sub
    
    chkAccNo(Index).Value = 1
    
    onPgm = True
    For i = 0 To 2
        If Index <> i Then
            chkAccNo(i).Value = 0
        End If
    Next i
    onPgm = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Save = True Then
        Query
    End If
End Sub

Private Sub Form_Activate()
'    medMain.lblSubMenu.Caption = Me.Caption

End Sub

Private Sub Form_Load()
    ClearAll
    Query
End Sub

Private Sub ClearAll()
    Dim i As Long
    
    onPgm = True
    For i = 0 To 2
        chkAccNo(i).Value = 0
    Next i
    onPgm = False
    
End Sub

Private Sub Query()
    Dim objcom003 As clsCom003
    Dim DrRS As Recordset
    Dim entdt As String
    Dim accno As Long
    Dim i As Long
    
    Set objcom003 = New clsCom003
    Set DrRS = objcom003.OpenRecordSet(BC2_ACCNO_CRITERION)
    If DrRS Is Nothing Then Exit Sub
    
    entdt = ""
    accno = -1
    lblAccNo = "-1"
    
    With DrRS
        If .RecordCount > 0 Then
            For i = 1 To .RecordCount
                If entdt = "" Then
                    entdt = .Fields("cdval1").Value & ""
                    accno = .Fields("field1").Value & ""
                    lblAccNo = accno
                Else
                    If entdt < .Fields("cdval1").Value & "" Then
                        entdt = .Fields("cdval1").Value & ""
                        accno = .Fields("field1").Value & ""
                        lblAccNo = accno
                    End If
                End If
                .MoveNext
            Next i
        End If
    End With
    
    Set DrRS = Nothing
    Set objcom003 = Nothing
    

    Select Case accno
        Case 0: chkAccNo(0).Value = 1
        Case 1: chkAccNo(1).Value = 1
        Case 2: chkAccNo(2).Value = 1
    End Select
End Sub

Private Function Save() As Boolean
    Dim objcom003 As clsCom003
    Dim accno As Long
    
    If chkAccNo(0).Value = 1 Then
        accno = 0
    ElseIf chkAccNo(1).Value = 1 Then
        accno = 1
    ElseIf chkAccNo(2).Value = 1 Then
        accno = 2
    Else
        accno = -1
    End If
    
    If accno < 0 Then
        MsgBox "접수번호 생성 기준을 선택 하십시요", vbCritical, Me.Caption
        Save = False
        Exit Function
    End If
    
    If accno = Val(lblAccNo) Then Exit Function
    
    Set objcom003 = New clsCom003
    objcom003.CDINDEX = BC2_ACCNO_CRITERION
    objcom003.cdval1 = Format(GetSystemDate, PRESENTDATE_FORMAT)
    objcom003.field1 = accno
    Save = objcom003.Save()
    Set objcom003 = Nothing
End Function




