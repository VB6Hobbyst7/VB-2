VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmReport 
   Caption         =   "STI-7 Real time Report"
   ClientHeight    =   15615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12720
   LinkTopic       =   "Form1"
   ScaleHeight     =   15615
   ScaleWidth      =   12720
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton cmdClose 
      Caption         =   "닫기"
      Height          =   855
      Left            =   11340
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "프린트"
      Height          =   855
      Left            =   11340
      TabIndex        =   1
      Top             =   90
      Width           =   1215
   End
   Begin FPSpread.vaSpread vasHPVReport 
      Height          =   15075
      Left            =   14280
      TabIndex        =   0
      Top             =   1770
      Width           =   10875
      _Version        =   393216
      _ExtentX        =   19182
      _ExtentY        =   26591
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
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      GridSolid       =   0   'False
      MaxCols         =   18
      MaxRows         =   43
      RetainSelBlock  =   0   'False
      ScrollBarMaxAlign=   0   'False
      ScrollBars      =   0
      ScrollBarShowMax=   0   'False
      SpreadDesigner  =   "frmReport.frx":0000
      UserResize      =   0
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   11895
      Left            =   13590
      TabIndex        =   3
      Top             =   690
      Visible         =   0   'False
      Width           =   9435
      _Version        =   393216
      _ExtentX        =   16642
      _ExtentY        =   20981
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
      GridSolid       =   0   'False
      MaxCols         =   11
      MaxRows         =   47
      RetainSelBlock  =   0   'False
      ScrollBarMaxAlign=   0   'False
      ScrollBars      =   0
      ScrollBarShowMax=   0   'False
      SpreadDesigner  =   "frmReport.frx":6084
      UserResize      =   0
   End
   Begin FPSpread.vaSpread vasSTI7Report 
      Height          =   15075
      Left            =   60
      TabIndex        =   4
      Top             =   90
      Width           =   10875
      _Version        =   393216
      _ExtentX        =   19182
      _ExtentY        =   26591
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
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      GridSolid       =   0   'False
      MaxCols         =   19
      MaxRows         =   40
      RetainSelBlock  =   0   'False
      ScrollBarMaxAlign=   0   'False
      ScrollBars      =   0
      ScrollBarShowMax=   0   'False
      SpreadDesigner  =   "frmReport.frx":9635
      UserResize      =   0
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub

Private Sub cmdPrint_Click()
    
    With vasSTI7Report
        .PrintOrientation = PrintOrientationPortrait 'PrintOrientationLandscape '가로출력
        .PrintColor = True
        .PrintType = PrintTypeAll
        .Action = 13
    End With
    
    Call cmdClose_Click
    
End Sub

Private Sub Form_Load()
    
    Call SetPrint_STI7

End Sub


Private Sub SetPrint()
    Dim i           As Integer
    Dim j           As Integer
    Dim varTmp      As Variant
    Dim varTmp2     As Variant
    Dim strTmp      As String
    Dim strINVALID  As String
    
    Dim strHData     As String
    Dim strLData     As String
    
    Dim varHNum     As Variant
    Dim varLNum     As Variant
    
    Dim blnHPos     As Boolean
    Dim blnLPos     As Boolean
    Dim blnMulti    As Boolean
    
    Dim strICData   As String
    Dim varICData   As Variant
    
    Dim varTemp1    As Variant
    Dim varTemp2    As Variant
    
'    blnPos = False
    blnMulti = False
    
    With vasHPVReport
        .Row = 5: .Col = 12: .Text = " " & varClipData(colJUBNO)                                '분자병리 접수번호
        .Row = 8: .Col = 14: .Text = " " & Format(varClipData(colRCPDATE), "####-##-##")                             '의뢰일자
        
        .Row = 10: .Col = 4: .Text = " " & varClipData(colCHARTNO)                              '등록번호
        .Row = 11: .Col = 4: .Text = " " & varClipData(colPNAME)                                '성     명
        .Row = 12: .Col = 4: .Text = " " & varClipData(colPSEX) & "/" & varClipData(colPAGE)    '성별/나이
        .Row = 13: .Col = 4: .Text = " " & varClipData(colPART) & "/" & varClipData(colROOM)    '과/병실

        .Row = 10: .Col = 14: .Text = " " & "Real-time PCR"                                     '검사 방법
        .Row = 11: .Col = 14: .Text = " " & varClipData(colSPCPART)                             '검체 종류
        .Row = 12: .Col = 14: .Text = " " & varClipData(colSPCNUM)                              '조직/세포번호
        .Row = 13: .Col = 14: .Text = " " & varClipData(colRELTEST)                             '상관성 검사결과
        
        strHData = ""
        strLData = ""
        
        varTmp = Split(varClipData(colITEMS), vbNewLine)
        
        For i = 0 To UBound(varTmp)
            If varTmp(i) <> "" Then
                If mGetP(varTmp(i), 1, ":") = "HPV High Risk Type " Then
                    If UCase(Trim(mGetP(mGetP(varTmp(i), 2, ":"), 1, "("))) = "POSITIVE" Then
                        strHData = "(" & mGetP(mGetP(varTmp(i), 2, ":"), 2, "(")
                        varHNum = Split(strHData, ",")
                        If UBound(varHNum) >= 1 Then
                            blnMulti = True
                        Else
                            blnMulti = False
                        End If
                        blnHPos = True
                    Else
                        strHData = "( )"
                    End If
                ElseIf mGetP(varTmp(i), 1, ":") = "HPV Low Risk Type " Then
                    If UCase(Trim(mGetP(mGetP(varTmp(i), 2, ":"), 1, "("))) = "POSITIVE" Then
                        strLData = "(" & mGetP(mGetP(varTmp(i), 2, ":"), 2, "(")
                        varLNum = Split(strLData, ",")
                        If UBound(varLNum) >= 1 Then
                            blnMulti = True
                        Else
                            blnMulti = False
                        End If
                        blnLPos = True
                    Else
                        strLData = "( )"
                    End If
                ElseIf Mid(varTmp(i), 1, 2) = "IC" Then
                    strICData = Mid(varTmp(i), 4)
                    If InStr(strICData, "/") > 0 Then
                        varICData = Split(strICData, "/")
                        If varICData(0) > varICData(1) Then
                            strICData = varICData(0)
                        Else
                            strICData = varICData(1)
                        End If
                    End If
                    .Row = 24: .Col = 16: .Text = "IC " & strICData            'IC Value
                    strINVALID = strICData
                    
                    '-- 2018.02.06 추가
                    If varClipData(colRESULT) = "INVALID" Then
                        strINVALID = "INVALID"
                    End If
                    
                    
                    Exit For
                End If
            End If
        Next
        
        If blnHPos = True And blnLPos = True Then
            blnMulti = True
        End If
        
        If blnMulti = True Then
            strHData = Replace(strHData, "(", "")
            strHData = Replace(strHData, ")", "")
            strLData = Replace(strLData, "(", "")
            strLData = Replace(strLData, ")", "")

            .Row = 24: .Col = 5: .Text = " □ "                                     'High-risk type ==> Check
            .Row = 25: .Col = 5: .Text = " □ "                                     'Low-risk type ==> Check
            .Row = 24: .Col = 6: .Text = "(  )"                                     'High-risk type ==> Value
            .Row = 25: .Col = 6: .Text = "(  )"                                     'Low-risk type ==> Value
            .Row = 26: .Col = 5: .Text = " ■ "                                     'Multiple infection  ==> Check
            .Row = 26: .Col = 6: .Text = "(" & strHData & "," & strLData & ")"      'Multiple infection  ==> Value
            
'            If strHData <> "" Then
'                .Row = 24: .Col = 6: .Text = strHData      'High-risk type ==> Value
'            Else
'                .Row = 24: .Col = 6: .Text = "(  )"        'High-risk type ==> Value
'            End If
'            If strLData <> "" Then
'                .Row = 25: .Col = 6: .Text = strLData      'Low-risk type ==> Value
'            Else
'                .Row = 25: .Col = 6: .Text = "(  )"        'Low-risk type ==> Value
'            End If
'
'            .Row = 26: .Col = 5: .Text = " ■ "                 'Multiple infection  ==> Check
'            strHData = Replace(strHData, "(", "")
'            strHData = Replace(strHData, ")", "")
'            strLData = Replace(strLData, "(", "")
'            strLData = Replace(strLData, ")", "")
'
'            .Row = 26: .Col = 6: .Text = "(" & strHData & "," & strLData & ")"    'Multiple infection  ==> Value
            
        Else
            If blnHPos = True Then
                .Row = 24: .Col = 5: .Text = " ■ "                                             'High-risk type ==> Check
                .Row = 24: .Col = 6: .Text = strHData                                           'High-risk type ==> Value
            Else
                .Row = 24: .Col = 5: .Text = " □ "                                             'High-risk type ==> Check
                .Row = 24: .Col = 6: .Text = "(  )"                                             'High-risk type ==> Value
            End If
            If blnLPos = True Then
                .Row = 25: .Col = 5: .Text = " ■ "                                             'Low-risk type ==> Check
                .Row = 25: .Col = 6: .Text = strLData                                           'Low-risk type ==> Value
            Else
                .Row = 25: .Col = 5: .Text = " □ "                                             'Low-risk type ==> Check
                .Row = 25: .Col = 6: .Text = "(  )"                                             'Low-risk type ==> Value
            End If
            .Row = 26: .Col = 5: .Text = " □ "                                                 'multi type ==> Check
            .Row = 26: .Col = 6: .Text = "(  )"                                                 'multi type ==> Value
            
        End If
        
        ' 초기화
        For i = 19 To 21 Step 2
            For j = 3 To 16
                .Row = i: .Col = j: .Text = "-"
            Next
        Next
        
'        Erase varTemp1
'        Erase varTemp2
        
        '-- High Set
        If strHData <> "( )" And strHData <> "" Then
            varTemp1 = Trim(strHData)
            varTemp1 = Replace(varTemp1, "(", "")
            varTemp1 = Replace(varTemp1, ")", "")
            varTemp1 = Trim(varTemp1)
            If varTemp1 <> "" Then
                varTemp1 = Split(varTemp1, ",")
                For i = 0 To UBound(varTemp1)
                    varTemp2 = Mid(varTemp1(i), 1, InStr(varTemp1(i), "+") - 1)
                    Select Case varTemp2
                        'A set
                        Case "66": .Row = 19: .Col = 3
                        Case "45": .Row = 19: .Col = 4
                        Case "58": .Row = 19: .Col = 5
                        Case "51": .Row = 19: .Col = 6
                        Case "59": .Row = 19: .Col = 7
                        Case "16": .Row = 19: .Col = 8
                        Case "33": .Row = 19: .Col = 9
                        Case "39": .Row = 19: .Col = 10
                        Case "52": .Row = 19: .Col = 11
                        Case "35": .Row = 19: .Col = 12
                        Case "18": .Row = 19: .Col = 13
                        Case "56": .Row = 19: .Col = 14
                        Case "68": .Row = 19: .Col = 15
                        Case "31": .Row = 19: .Col = 16
                        'B set
                        Case "26": .Row = 21: .Col = 3
                        Case "69": .Row = 21: .Col = 4
                        Case "73": .Row = 21: .Col = 5
                        Case "42": .Row = 21: .Col = 6
                        Case "82": .Row = 21: .Col = 7
                        Case "53": .Row = 21: .Col = 8
                        Case "43": .Row = 21: .Col = 9
                        Case "54": .Row = 21: .Col = 10
                        Case "70": .Row = 21: .Col = 11
                        Case "61": .Row = 21: .Col = 12
                        Case "6":  .Row = 21: .Col = 13
                        Case "44": .Row = 21: .Col = 14
                        Case "40": .Row = 21: .Col = 15
                        Case "11": .Row = 21: .Col = 16
                        
                    End Select
                    .Text = Mid(varTemp1(i), InStr(varTemp1(i), "+"))
                Next
            End If
        End If
        
        '-- Low set
        If strLData <> "( )" And strLData <> "" Then
            varTemp1 = Trim(strLData)
            varTemp1 = Replace(varTemp1, "(", "")
            varTemp1 = Replace(varTemp1, ")", "")
            varTemp1 = Trim(varTemp1)
            If varTemp1 <> "" Then
                varTemp1 = Split(varTemp1, ",")
                For i = 0 To UBound(varTemp1)
                    varTemp2 = Mid(varTemp1(i), 1, InStr(varTemp1(i), "+") - 1)
                    Select Case varTemp2
                        'A set
                        Case "66": .Row = 19: .Col = 3
                        Case "45": .Row = 19: .Col = 4
                        Case "58": .Row = 19: .Col = 5
                        Case "51": .Row = 19: .Col = 6
                        Case "59": .Row = 19: .Col = 7
                        Case "16": .Row = 19: .Col = 8
                        Case "33": .Row = 19: .Col = 9
                        Case "39": .Row = 19: .Col = 10
                        Case "52": .Row = 19: .Col = 11
                        Case "35": .Row = 19: .Col = 12
                        Case "18": .Row = 19: .Col = 13
                        Case "56": .Row = 19: .Col = 14
                        Case "68": .Row = 19: .Col = 15
                        Case "31": .Row = 19: .Col = 16
                        'B set
                        Case "26": .Row = 21: .Col = 3
                        Case "69": .Row = 21: .Col = 4
                        Case "73": .Row = 21: .Col = 5
                        Case "42": .Row = 21: .Col = 6
                        Case "82": .Row = 21: .Col = 7
                        Case "53": .Row = 21: .Col = 8
                        Case "43": .Row = 21: .Col = 9
                        Case "54": .Row = 21: .Col = 10
                        Case "70": .Row = 21: .Col = 11
                        Case "61": .Row = 21: .Col = 12
                        Case "6":  .Row = 21: .Col = 13
                        Case "44": .Row = 21: .Col = 14
                        Case "40": .Row = 21: .Col = 15
                        Case "11": .Row = 21: .Col = 16
                        
                    End Select
                    .Text = Mid(varTemp1(i), InStr(varTemp1(i), "+"))
                Next
            End If
        End If
        
        .Row = 19: .Col = 17: .Text = varICData(0)
        .Row = 21: .Col = 17: .Text = varICData(1)
        
                
        If blnHPos = True Or blnLPos = True Then
            .Row = 24: .Col = 15: .Text = " □ "        'Negative 여부
        Else
            .Row = 24: .Col = 15: .Text = " ■ "        'Negative 여부
        End If
        
        
        If UCase(Trim(strINVALID)) = "INVALID" Then
            .Row = 25: .Col = 16: .Text = "Invalid"         'Invalid
        Else
            .Row = 25: .Col = 16: .Text = ""                'Invalid
        End If
                      
'        .Row = 27: .Col = 2: .Text = " * 상관성 검사결과 : " & varClipData(colRELTEST) '상관성 검사결과

        .Row = 39: .Col = 13:   .Text = "M" & Format(Now, "yy") & "-"
        
        .Row = 42: .Col = 5:    .Text = varClipData(colRSTDATE)  '검사보고일 'Format(Now, "yyyy년 mm월 dd일")
        .Row = 42: .Col = 15:   .Text = Trim(mGetP(varClipData(colDOCTOR), 2, "-")) '판독의사   'APD16 - 김한성 M.D.
        
    End With
End Sub

Private Sub SetPrint_STI7()
    Dim i           As Integer
    Dim j           As Integer
    Dim varTmp      As Variant
    Dim varTmp2     As Variant
    Dim strTmp      As String
    Dim strINVALID  As String
    
    Dim strHData     As String
    Dim strLData     As String
    
    Dim varHNum     As Variant
    Dim varLNum     As Variant
    
    Dim blnHPos     As Boolean
    Dim blnLPos     As Boolean
    Dim blnMulti    As Boolean
    
    Dim strICData   As String
    Dim varICData   As Variant
    
    Dim varTemp1    As Variant
    Dim varTemp2    As Variant
    
    Dim strUU       As String
    Dim strUP       As String
    Dim strMG       As String
    Dim strMH       As String
    Dim strNG       As String
    Dim strCT       As String
    Dim strTV       As String
    Dim strIC       As String
    
    Dim strABData    As String
    Dim strSData     As String
    Dim strMData     As String
    
    strUU = ""
    strUP = ""
    strMG = ""
    strMH = ""
    strNG = ""
    strCT = ""
    strTV = ""
    strIC = ""
    
    strICData = ""
    strSData = ""
    strMData = ""
    strABData = ""
    strINVALID = ""
    
    blnMulti = False
    
    With vasSTI7Report
        .Row = 5: .Col = 12: .Text = " " & varClipData(colJUBNO)                                '분자병리 접수번호
        .Row = 8: .Col = 14: .Text = " " & Format(varClipData(colRCPDATE), "####-##-##")                             '의뢰일자
        
        .Row = 10: .Col = 4: .Text = " " & varClipData(colCHARTNO)                              '등록번호
        .Row = 11: .Col = 4: .Text = " " & varClipData(colPNAME)                                '성     명
        .Row = 12: .Col = 4: .Text = " " & varClipData(colPSEX) & "/" & varClipData(colPAGE)    '성별/나이
        .Row = 13: .Col = 4: .Text = " " & varClipData(colPART) & "/" & varClipData(colROOM)    '과/병실

        .Row = 10: .Col = 14: .Text = " " & "Real-time PCR"                                     '검사 방법
        .Row = 11: .Col = 14: .Text = " " & varClipData(colSPCPART)                             '검체 종류
        .Row = 12: .Col = 14: .Text = " " & varClipData(colSPCNUM)                              '조직/세포번호
        .Row = 13: .Col = 14: .Text = " " & varClipData(colRELTEST)                             '상관성 검사결과
        
        strHData = ""
        strLData = ""
        
        varTmp = Split(varClipData(colITEMS), vbNewLine)
        
        For i = 0 To UBound(varTmp)
            If i = 0 And varTmp(i) <> "" Then
                If UCase(Trim(varTmp(i))) = "INVALID" Then
                    strINVALID = "Invalid"
                Else
                    strINVALID = ""
                    If UCase(Trim(varTmp(i))) = "NEGATIVE" Then
                        strABData = "0"
                    Else
                        strABData = "1"
                    End If
                    
                    If strABData = "0" Then
                        .Row = 22: .Col = 15: .Text = " ■ "        'Negative 여부
                    Else
                        .Row = 22: .Col = 15: .Text = " □ "        'Negative 여부
                    End If
                End If
                
            ElseIf i = 1 And varTmp(i) <> "" Then
                varTmp2 = Split(varTmp(i), ",")
                For j = 0 To UBound(varTmp2)
                    Select Case mGetP(varTmp2(j), 1, ":")
                        Case "UU": strUU = mGetP(varTmp2(j), 2, ":"): .Row = 19: .Col = 3:  .Text = strUU     'UU
                        Case "UP": strUP = mGetP(varTmp2(j), 2, ":"): .Row = 19: .Col = 5:  .Text = strUP     'UP
                        Case "MG": strMG = mGetP(varTmp2(j), 2, ":"): .Row = 19: .Col = 7:  .Text = strMG     'MG
                        Case "MH": strMH = mGetP(varTmp2(j), 2, ":"): .Row = 19: .Col = 9:  .Text = strMH     'MH
                        Case "NG": strNG = mGetP(varTmp2(j), 2, ":"): .Row = 19: .Col = 11: .Text = strNG     'NG
                        Case "CT": strCT = mGetP(varTmp2(j), 2, ":"): .Row = 19: .Col = 13: .Text = strCT     'CT
                        Case "TV": strTV = mGetP(varTmp2(j), 2, ":"): .Row = 19: .Col = 15: .Text = strTV     'TV
                    End Select
                Next
            
            ElseIf i = 2 And varTmp(i) <> "" Then
                'Single Infection
                If Trim(mGetP(varTmp(i), 2, ":")) = "N/A" Then
                    .Row = 23: .Col = 5: .Text = " □ "
                    .Row = 23: .Col = 6: .Text = ""
                Else
                    strSData = Replace(varTmp(i), mGetP(varTmp(i), 1, ":"), "")
                    'strSData = Mid(Trim(strSData), 2)
                    strSData = Replace(strSData, ":", "")
                    
                    .Row = 23: .Col = 5: .Text = " ■ "
                    .Row = 23: .Col = 6: .Text = strSData
                End If
            
            ElseIf i = 3 And varTmp(i) <> "" Then
                'Multi Infection
                If Trim(mGetP(varTmp(i), 2, ":")) = "N/A" Then
                    .Row = 24: .Col = 5: .Text = " □ "
                    .Row = 24: .Col = 6: .Text = ""
                Else
                    strMData = Replace(varTmp(i), mGetP(varTmp(i), 1, ":"), "")
                    strMData = Replace(strMData, ":", "")
                    
                    .Row = 24: .Col = 5: .Text = " ■ "
                    .Row = 24: .Col = 6: .Text = strMData
                End If
            ElseIf i = 4 And varTmp(i) <> "" Then
                'IC
                If strINVALID = "Invalid" Then
                    .Row = 24: .Col = 15: .Text = " ■ "
                    .Row = 24: .Col = 16: .Text = "Invalid"
                Else
                    .Row = 24: .Col = 15: .Text = " □ "
                    .Row = 24: .Col = 16: .Text = ""
                End If
                strIC = mGetP(varTmp(i), 2, ":")
                varTmp(i) = Replace(varTmp(i), ":", " ")
                .Row = 22: .Col = 16: .Text = varTmp(i)
                .Row = 19: .Col = 17: .Text = strIC
            End If
        Next
        
'        If blnHPos = True And blnLPos = True Then
'            blnMulti = True
'        End If
'
'        If blnMulti = True Then
'            strHData = Replace(strHData, "(", "")
'            strHData = Replace(strHData, ")", "")
'            strLData = Replace(strLData, "(", "")
'            strLData = Replace(strLData, ")", "")
'
'            .Row = 24: .Col = 5: .Text = " □ "                                     'High-risk type ==> Check
'            .Row = 25: .Col = 5: .Text = " □ "                                     'Low-risk type ==> Check
'            .Row = 24: .Col = 6: .Text = "(  )"                                     'High-risk type ==> Value
'            .Row = 25: .Col = 6: .Text = "(  )"                                     'Low-risk type ==> Value
'            .Row = 26: .Col = 5: .Text = " ■ "                                     'Multiple infection  ==> Check
'            .Row = 26: .Col = 6: .Text = "(" & strHData & "," & strLData & ")"      'Multiple infection  ==> Value
'
''            If strHData <> "" Then
''                .Row = 24: .Col = 6: .Text = strHData      'High-risk type ==> Value
''            Else
''                .Row = 24: .Col = 6: .Text = "(  )"        'High-risk type ==> Value
''            End If
''            If strLData <> "" Then
''                .Row = 25: .Col = 6: .Text = strLData      'Low-risk type ==> Value
''            Else
''                .Row = 25: .Col = 6: .Text = "(  )"        'Low-risk type ==> Value
''            End If
''
''            .Row = 26: .Col = 5: .Text = " ■ "                 'Multiple infection  ==> Check
''            strHData = Replace(strHData, "(", "")
''            strHData = Replace(strHData, ")", "")
''            strLData = Replace(strLData, "(", "")
''            strLData = Replace(strLData, ")", "")
''
''            .Row = 26: .Col = 6: .Text = "(" & strHData & "," & strLData & ")"    'Multiple infection  ==> Value
'
'        Else
'            If blnHPos = True Then
'                .Row = 24: .Col = 5: .Text = " ■ "                                             'High-risk type ==> Check
'                .Row = 24: .Col = 6: .Text = strHData                                           'High-risk type ==> Value
'            Else
'                .Row = 24: .Col = 5: .Text = " □ "                                             'High-risk type ==> Check
'                .Row = 24: .Col = 6: .Text = "(  )"                                             'High-risk type ==> Value
'            End If
'            If blnLPos = True Then
'                .Row = 25: .Col = 5: .Text = " ■ "                                             'Low-risk type ==> Check
'                .Row = 25: .Col = 6: .Text = strLData                                           'Low-risk type ==> Value
'            Else
'                .Row = 25: .Col = 5: .Text = " □ "                                             'Low-risk type ==> Check
'                .Row = 25: .Col = 6: .Text = "(  )"                                             'Low-risk type ==> Value
'            End If
'            .Row = 26: .Col = 5: .Text = " □ "                                                 'multi type ==> Check
'            .Row = 26: .Col = 6: .Text = "(  )"                                                 'multi type ==> Value
'
'        End If
'
'        ' 초기화
'        For i = 19 To 21 Step 2
'            For j = 3 To 16
'                .Row = i: .Col = j: .Text = "-"
'            Next
'        Next
'
''        Erase varTemp1
''        Erase varTemp2
'
'        '-- High Set
'        If strHData <> "( )" And strHData <> "" Then
'            varTemp1 = Trim(strHData)
'            varTemp1 = Replace(varTemp1, "(", "")
'            varTemp1 = Replace(varTemp1, ")", "")
'            varTemp1 = Trim(varTemp1)
'            If varTemp1 <> "" Then
'                varTemp1 = Split(varTemp1, ",")
'                For i = 0 To UBound(varTemp1)
'                    varTemp2 = Mid(varTemp1(i), 1, InStr(varTemp1(i), "+") - 1)
'                    Select Case varTemp2
'                        'A set
'                        Case "66": .Row = 19: .Col = 3
'                        Case "45": .Row = 19: .Col = 4
'                        Case "58": .Row = 19: .Col = 5
'                        Case "51": .Row = 19: .Col = 6
'                        Case "59": .Row = 19: .Col = 7
'                        Case "16": .Row = 19: .Col = 8
'                        Case "33": .Row = 19: .Col = 9
'                        Case "39": .Row = 19: .Col = 10
'                        Case "52": .Row = 19: .Col = 11
'                        Case "35": .Row = 19: .Col = 12
'                        Case "18": .Row = 19: .Col = 13
'                        Case "56": .Row = 19: .Col = 14
'                        Case "68": .Row = 19: .Col = 15
'                        Case "31": .Row = 19: .Col = 16
'                        'B set
'                        Case "26": .Row = 21: .Col = 3
'                        Case "69": .Row = 21: .Col = 4
'                        Case "73": .Row = 21: .Col = 5
'                        Case "42": .Row = 21: .Col = 6
'                        Case "82": .Row = 21: .Col = 7
'                        Case "53": .Row = 21: .Col = 8
'                        Case "43": .Row = 21: .Col = 9
'                        Case "54": .Row = 21: .Col = 10
'                        Case "70": .Row = 21: .Col = 11
'                        Case "61": .Row = 21: .Col = 12
'                        Case "6":  .Row = 21: .Col = 13
'                        Case "44": .Row = 21: .Col = 14
'                        Case "40": .Row = 21: .Col = 15
'                        Case "11": .Row = 21: .Col = 16
'
'                    End Select
'                    .Text = Mid(varTemp1(i), InStr(varTemp1(i), "+"))
'                Next
'            End If
'        End If
'
'        '-- Low set
'        If strLData <> "( )" And strLData <> "" Then
'            varTemp1 = Trim(strLData)
'            varTemp1 = Replace(varTemp1, "(", "")
'            varTemp1 = Replace(varTemp1, ")", "")
'            varTemp1 = Trim(varTemp1)
'            If varTemp1 <> "" Then
'                varTemp1 = Split(varTemp1, ",")
'                For i = 0 To UBound(varTemp1)
'                    varTemp2 = Mid(varTemp1(i), 1, InStr(varTemp1(i), "+") - 1)
'                    Select Case varTemp2
'                        'A set
'                        Case "66": .Row = 19: .Col = 3
'                        Case "45": .Row = 19: .Col = 4
'                        Case "58": .Row = 19: .Col = 5
'                        Case "51": .Row = 19: .Col = 6
'                        Case "59": .Row = 19: .Col = 7
'                        Case "16": .Row = 19: .Col = 8
'                        Case "33": .Row = 19: .Col = 9
'                        Case "39": .Row = 19: .Col = 10
'                        Case "52": .Row = 19: .Col = 11
'                        Case "35": .Row = 19: .Col = 12
'                        Case "18": .Row = 19: .Col = 13
'                        Case "56": .Row = 19: .Col = 14
'                        Case "68": .Row = 19: .Col = 15
'                        Case "31": .Row = 19: .Col = 16
'                        'B set
'                        Case "26": .Row = 21: .Col = 3
'                        Case "69": .Row = 21: .Col = 4
'                        Case "73": .Row = 21: .Col = 5
'                        Case "42": .Row = 21: .Col = 6
'                        Case "82": .Row = 21: .Col = 7
'                        Case "53": .Row = 21: .Col = 8
'                        Case "43": .Row = 21: .Col = 9
'                        Case "54": .Row = 21: .Col = 10
'                        Case "70": .Row = 21: .Col = 11
'                        Case "61": .Row = 21: .Col = 12
'                        Case "6":  .Row = 21: .Col = 13
'                        Case "44": .Row = 21: .Col = 14
'                        Case "40": .Row = 21: .Col = 15
'                        Case "11": .Row = 21: .Col = 16
'
'                    End Select
'                    .Text = Mid(varTemp1(i), InStr(varTemp1(i), "+"))
'                Next
'            End If
'        End If
'
'        .Row = 19: .Col = 17: .Text = varICData(0)
'        .Row = 21: .Col = 17: .Text = varICData(1)
'
'
'        If blnHPos = True Or blnLPos = True Then
'            .Row = 24: .Col = 15: .Text = " □ "        'Negative 여부
'        Else
'            .Row = 24: .Col = 15: .Text = " ■ "        'Negative 여부
'        End If
'
'
'        If UCase(Trim(strINVALID)) = "INVALID" Then
'            .Row = 25: .Col = 16: .Text = "Invalid"         'Invalid
'        Else
'            .Row = 25: .Col = 16: .Text = ""                'Invalid
'        End If
                      
'        .Row = 27: .Col = 2: .Text = " * 상관성 검사결과 : " & varClipData(colRELTEST) '상관성 검사결과

        '.Row = 39: .Col = 13:   .Text = "M" & Format(Now, "yy") & "-"
        
        .Row = 39: .Col = 5:    .Text = varClipData(colRSTDATE)  '검사보고일 'Format(Now, "yyyy년 mm월 dd일")
        .Row = 39: .Col = 15:   .Text = Trim(mGetP(varClipData(colDOCTOR), 2, "-")) '판독의사   'APD16 - 김한성 M.D.
        
        'Note
        .Row = 37: .Col = 2:    .Text = varClipData(colREMARK)  'Note
        
    End With
End Sub

