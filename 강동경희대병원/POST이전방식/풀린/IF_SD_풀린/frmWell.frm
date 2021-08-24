VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmWell 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  '단일 고정
   Caption         =   "Well 지정"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12915
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   12915
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame fraWell 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Well 지정"
      Height          =   3945
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   12735
      Begin VB.CommandButton cmdCLR 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clear"
         Height          =   315
         Left            =   5700
         Style           =   1  '그래픽
         TabIndex        =   12
         Top             =   210
         Width           =   915
      End
      Begin VB.TextBox txtWellCnt 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Text            =   "5"
         Top             =   240
         Width           =   555
      End
      Begin VB.CheckBox chkSingle 
         BackColor       =   &H00FFFFFF&
         Caption         =   "단독검체"
         Height          =   255
         Left            =   1620
         TabIndex        =   10
         Top             =   270
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   1575
         Left            =   1620
         TabIndex        =   4
         Top             =   1470
         Visible         =   0   'False
         Width           =   3225
         Begin VB.Timer tmrAuto 
            Left            =   2160
            Top             =   210
         End
         Begin VB.TextBox txtRow 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   1410
            Locked          =   -1  'True
            TabIndex        =   8
            Text            =   "5"
            Top             =   270
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.CommandButton cmdHidden 
            BackColor       =   &H00FFFFFF&
            Caption         =   "X"
            Height          =   285
            Left            =   1800
            Style           =   1  '그래픽
            TabIndex        =   7
            Top             =   810
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.CheckBox chkHori 
            BackColor       =   &H00FFFFFF&
            Caption         =   "세로"
            Height          =   285
            Left            =   1800
            TabIndex        =   6
            Top             =   1140
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CommandButton cmdAuto 
            BackColor       =   &H00FFFFFF&
            Caption         =   "자동발번"
            Height          =   315
            Left            =   420
            Style           =   1  '그래픽
            TabIndex        =   5
            Top             =   1110
            Width           =   915
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "현재열 : "
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   3
            Left            =   660
            TabIndex        =   9
            Top             =   330
            Visible         =   0   'False
            Width           =   660
            WordWrap        =   -1  'True
         End
      End
      Begin VB.TextBox txtTestCnt 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3570
         TabIndex        =   3
         Text            =   "96"
         Top             =   240
         Width           =   555
      End
      Begin VB.CheckBox chkAuto 
         BackColor       =   &H00FFFFFF&
         Caption         =   "자동발번"
         Height          =   285
         Left            =   4200
         TabIndex        =   2
         Top             =   240
         Value           =   1  '확인
         Visible         =   0   'False
         Width           =   1065
      End
      Begin FPSpread.vaSpread spdWell 
         Height          =   3315
         Left            =   120
         TabIndex        =   0
         Top             =   570
         Width           =   12465
         _Version        =   393216
         _ExtentX        =   21987
         _ExtentY        =   5847
         _StockProps     =   64
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   24
         MaxRows         =   8
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ScrollBars      =   0
         ScrollBarShowMax=   0   'False
         SpreadDesigner  =   "frmWell.frx":0000
         UserResize      =   0
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "검사수 : "
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   4
         Left            =   180
         TabIndex        =   14
         Top             =   300
         Width           =   690
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "검사수 : "
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   2820
         TabIndex        =   13
         Top             =   300
         Width           =   780
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmWell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngLastRow As Long
Dim lngLastCol As Long
Dim lngTotCnt  As Long


Private Sub chkSingle_Click()
    
    If chkSingle.Value = "1" Then
        txtWellCnt.Text = "1"
        'chkAuto.Value = "1"
    Else
        txtWellCnt.Text = "4"
        'chkAuto.Value = "0"
    End If

End Sub

Private Sub cmdCLR_Click()
    Dim intRow, intCol As Integer
    
    With spdWell
        For intRow = 1 To .MaxRows
            For intCol = 1 To .MaxCols
                .Row = intRow
                .Col = intCol
                .BackColor = vbWhite
            Next
        Next
    End With

    With frmMain.spdOrder
        For intRow = 1 To .MaxRows
            Call SetText(frmMain.spdOrder, "", intRow, colRACKNO)
            Call SetText(frmMain.spdOrder, "", intRow, colPOSNO)
        Next
    End With

    txtRow.Text = 1
    lngTotCnt = 0
    
End Sub


Private Sub cmdHidden_Click()
    
    'fraWell.Visible = False
    Unload Me
    
End Sub

Private Sub Form_Load()
    Dim intRow  As Integer
    Dim intCol  As Integer
    Dim strRow  As String
    Dim strCol  As String
    Dim intNum  As Integer
    Dim strTmp  As String
    
    intNum = 0
    strTmp = ""
'    With spdWell
'        For intRow = 1 To .MaxRows
'            For intCol = 1 To .MaxCols
'                Select Case intRow
'                    Case 1: strRow = "A"
'                    Case 2: strRow = "B"
'                    Case 3: strRow = "C"
'                    Case 4: strRow = "D"
'                    Case 5: strRow = "E"
'                    Case 6: strRow = "F"
'                    Case 7: strRow = "G"
'                    Case 8: strRow = "H"
'                End Select
'
'                Select Case intCol
'                    Case 1 To 4:    strCol = "01"
'                    Case 5 To 8:    strCol = "02"
'                    Case 9 To 12:   strCol = "03"
'                End Select
'                'Call SetText(spdWell, strRow & Format(intCol, "00"), intRow, intCol)
'                Call SetTag(spdWell, strRow & strCol, intRow, intCol)
'            Next
'        Next
'    End With
'
'    With spdWell
'        For intCol = 1 To .MaxCols
'            For intRow = 1 To .MaxRows
'                intNum = intNum + 1
'                strTmp = GetTag(spdWell, intRow, intCol)
'                strTmp = Space(4) & intNum & Space(5) & "/" & strTmp
'                Call SetText(spdWell, strTmp, intRow, intCol)
'            Next
'        Next
'    End With
    
    With spdWell
        For intRow = 1 To .MaxRows
            For intCol = 1 To .MaxCols
                Select Case intRow
                    Case 1:     strRow = "A"
                    Case 2:     strRow = "B"
                    Case 3:     strRow = "C"
                    Case 4:     strRow = "D"
                    Case 5:     strRow = "E"
                    Case 6:     strRow = "F"
                    Case 7:     strRow = "G"
                    Case 8:     strRow = "H"
                    Case 9:     strRow = "I"
                    Case 10:    strRow = "J"
                    Case 11:    strRow = "K"
                    Case 12:    strRow = "L"
                    Case 13:    strRow = "M"
                    Case 14:    strRow = "N"
                    Case 15:    strRow = "O"
                    Case 16:    strRow = "P"
                End Select
                
'                Select Case intCol
'                    Case 1 To 4:    strCol = "01"
'                    Case 5 To 8:    strCol = "02"
'                    Case 9 To 12:   strCol = "03"
'                    Case 13 To 16:  strCol = "04"
'                    Case 17 To 20:  strCol = "05"
'                    Case 21 To 24:  strCol = "06"
'                End Select
                Call SetText(spdWell, strRow & Format(intCol, "00"), intRow, intCol)
                Call SetTag(spdWell, strRow & Format(intCol, "00"), intRow, intCol)
            Next
        Next
    End With
    
    With spdWell
        For intCol = 1 To .MaxCols
            For intRow = 1 To .MaxRows
                intNum = intNum + 1
                strTmp = GetTag(spdWell, intRow, intCol)
                strTmp = Space(4) & intNum & Space(2) & "/" & Space(2) & strTmp
                Call SetText(spdWell, strTmp, intRow, intCol)
            Next
        Next
    End With
    
    tmrAuto.Enabled = False

    lngTotCnt = 0

End Sub

Private Sub spdWell_Click(ByVal Col As Long, ByVal Row As Long)
    Dim intRow  As Integer
    Dim intCol  As Integer
    Dim intRow2 As Integer
    Dim intCol2 As Integer
    Dim i, J As Integer
    Dim intR As Integer
    Dim intG As Integer
    Dim intB As Integer
    Dim strTmp      As String
    Dim strWell     As String
    Dim strSeqNo    As String
    Dim intWellCnt  As Integer
    
'''    If chkHori.Value = "0" Then
'''        intRow2 = 0
'''        intCol2 = 0
'''        intRow = Row
'''        intCol = Col
'''
'''        With spdWell
'''            For i = intCol To (intCol + CCur(txtWellCnt.Text)) - 1
'''                If intCol2 = 0 Then
'''                    .Col = i
'''                Else
'''                    intCol2 = intCol2 + 1
'''                    .Col = intCol2
'''                End If
'''
'''                If i > 24 Then
'''                    If intRow2 = 0 Then
'''                        intRow = intRow + 1
'''                        intRow2 = 1
'''                        intCol2 = 1
'''                    End If
'''                    .Col = intCol2
'''                End If
'''                .Row = intRow
'''
'''                strTmp = .Text
'''                strSeqNo = Trim(mGetP(strTmp, 1, "/"))
'''                If strWell = "" Then
'''                    strWell = Trim(mGetP(strTmp, 2, "/"))
'''                End If
'''
'''                If .BackColor = vbWhite Then
'''                    .BackColor = vbYellow
'''                    DoEvents
'''                    '메인화면에서 같은 SEQ 찾기
'''                    For J = 1 To frmMain.spdOrder.MaxRows
'''                        frmMain.spdOrder.Row = J
'''                        frmMain.spdOrder.Col = colSEQNO
'''                        If frmMain.spdOrder.Text = strSeqNo Then
'''                            Call SetText(frmMain.spdOrder, Mid(strWell, 1, 1), J, colRACKNO)
'''                            Call SetText(frmMain.spdOrder, Mid(strWell, 2), J, colPOSNO)
'''
'''                            lngTotCnt = lngTotCnt + 1
'''                            If lngTotCnt >= txtTestCnt.Text Then
'''                                tmrAuto.Enabled = False
'''                                Exit Sub
'''                            End If
'''                            Exit For
'''                        End If
'''                    Next
'''                Else
'''                    .BackColor = vbWhite
'''                End If
'''            Next
'''
'''            If Row = 8 Then
'''                lngLastRow = 0
'''                lngLastCol = Col + 4
'''            Else
'''                lngLastRow = .Row
'''                lngLastCol = Col
'''            End If
'''            DoEvents
'''
'''            If lngTotCnt >= txtTestCnt.Text Then
'''                tmrAuto.Enabled = False
'''                Exit Sub
'''            End If
'''            If chkAuto.Value = "1" Then
'''                tmrAuto.Interval = 100
'''                tmrAuto.Enabled = True
'''            End If
'''
'''        End With
'''    Else
'''        intRow2 = 0
'''        intCol2 = 0
'''        intRow = Row
'''        intCol = Col
'''
'''        With spdWell
'''            For i = intRow To (intRow + CCur(txtWellCnt.Text)) - 1
'''                If i > 8 Then
'''                    If intCol2 = 0 Then
'''                        intCol = intCol + 1
'''                        intRow2 = 1
'''                        intCol2 = 1
'''                    End If
'''                End If
'''                .Col = intCol
'''
'''                If intRow2 = 0 Then
'''                    .Row = i
'''                Else
'''                    .Row = intRow2
'''                     If intRow2 >= 1 Then
'''                         intRow2 = intRow2 + 1
'''                     End If
'''                End If
'''
'''                If .BackColor = vbWhite Then
'''                    .BackColor = vbYellow
'''                    Call SetText(frmMain.spdOrder, Mid(spdWell.Text, 1, 1), txtRow.Text, colRACKNO)
'''                    Call SetText(frmMain.spdOrder, Mid(spdWell.Text, 2), txtRow.Text, colPOSNO)
'''                    txtRow.Text = txtRow.Text + 1
'''                Else
'''                    .BackColor = vbWhite
'''                End If
'''            Next
'''            lngLastRow = .Row
'''            lngLastCol = .Col
'''            DoEvents
'''        End With
'''    End If

    
    intWellCnt = 0
    
    If chkHori.Value = "0" Then
        intRow2 = 0
        intCol2 = 0
        intRow = Row
        intCol = Col
    
        With spdWell
            For i = Col To .MaxCols
                For J = Row To .MaxRows
                    strTmp = GetText(spdWell, J, i)
                    strSeqNo = Trim(mGetP(strTmp, 1, "/"))
                    If strWell = "" Then
                        strWell = Trim(mGetP(strTmp, 2, "/"))
                        mResult.RackNo = Mid(strWell, 1, 1)
                        mResult.TubePos = Mid(strWell, 2)
                        intWellCnt = 0
                    End If
                    
                    .Row = J
                    .Col = i
                    If .BackColor = vbWhite Then
                        .BackColor = vbYellow
                        DoEvents
                        '메인화면에서 같은 SEQ 찾기
                        For intRow = 1 To frmMain.spdOrder.MaxRows
                            frmMain.spdOrder.Row = intRow
                            frmMain.spdOrder.Col = colSEQNO
                            If frmMain.spdOrder.Text = strSeqNo Then
                                Call SetText(frmMain.spdOrder, mResult.RackNo, intRow, colRACKNO)
                                Call SetText(frmMain.spdOrder, mResult.TubePos, intRow, colPOSNO)
                                
                                intWellCnt = intWellCnt + 1
                                lngTotCnt = lngTotCnt + 1
                                
                                If lngTotCnt = 5 Then
                                    mResult.TubePos = CCur(mResult.TubePos) + 1
                                    If CCur(mResult.TubePos) = 9 Then
                                        mResult.TubePos = "01"
                                        'mResult.RackNo = mResult.RackNo + 1
                                        mResult.RackNo = Chr(Asc(mResult.RackNo) + 1)
                                    End If
                                    
                                    mResult.TubePos = Format(mResult.TubePos, "00")
                                    lngTotCnt = 0
                                End If
                                
                                If lngTotCnt >= txtTestCnt.Text Then
                                    'tmrAuto.Enabled = False
                                    Exit Sub
                                End If
                                Exit For
                            End If
                        Next
                    Else
                        .BackColor = vbWhite
                    End If
                Next
            Next
            
'            If Row = 8 Then
'                lngLastRow = 0
'                lngLastCol = Col + 4
'            Else
'                lngLastRow = .Row
'                lngLastCol = Col
'            End If
'            DoEvents
'
'            If lngTotCnt >= txtTestCnt.Text Then
'                tmrAuto.Enabled = False
'                Exit Sub
'            End If
'            If chkAuto.Value = "1" Then
'                tmrAuto.Interval = 100
'                tmrAuto.Enabled = True
'            End If
            
        End With
    Else
        intRow2 = 0
        intCol2 = 0
        intRow = Row
        intCol = Col
    
        With spdWell
            For i = intRow To (intRow + CCur(txtWellCnt.Text)) - 1
                If i > 8 Then
                    If intCol2 = 0 Then
                        intCol = intCol + 1
                        intRow2 = 1
                        intCol2 = 1
                    End If
                End If
                .Col = intCol
    
                If intRow2 = 0 Then
                    .Row = i
                Else
                    .Row = intRow2
                     If intRow2 >= 1 Then
                         intRow2 = intRow2 + 1
                     End If
                End If
    
                If .BackColor = vbWhite Then
                    .BackColor = vbYellow
                    Call SetText(frmMain.spdOrder, Mid(spdWell.Text, 1, 1), txtRow.Text, colRACKNO)
                    Call SetText(frmMain.spdOrder, Mid(spdWell.Text, 2), txtRow.Text, colPOSNO)
                    txtRow.Text = txtRow.Text + 1
                Else
                    .BackColor = vbWhite
                End If
            Next
            lngLastRow = .Row
            lngLastCol = .Col
            DoEvents
        End With
    End If
    
End Sub

Private Sub tmrAuto_Timer()

    Call spdWell_Click(lngLastCol, lngLastRow + 1)

    If GetText(frmMain.spdOrder, frmMain.spdOrder.MaxRows, colRACKNO) <> "" Then
        tmrAuto.Enabled = False
    End If

End Sub

