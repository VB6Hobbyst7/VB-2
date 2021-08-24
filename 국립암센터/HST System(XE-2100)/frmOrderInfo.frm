VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Begin VB.Form frmOrderInfo 
   Caption         =   "WorkList 조회 및 오더 전송"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10425
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOrderInfo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   10425
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdSet 
      Caption         =   "SET"
      Height          =   315
      Left            =   3660
      TabIndex        =   23
      Top             =   780
      Width           =   1005
   End
   Begin VB.TextBox txtInt 
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1740
      TabIndex        =   20
      Text            =   "10"
      Top             =   810
      Width           =   615
   End
   Begin VB.CommandButton cmdPortOpen 
      Caption         =   "PortOpen"
      Height          =   495
      Left            =   2580
      TabIndex        =   18
      Top             =   1230
      Width           =   1365
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5340
      TabIndex        =   17
      Top             =   1230
      Width           =   1365
   End
   Begin VB.CheckBox chkMode 
      Caption         =   "AUTO"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   270
      Style           =   1  '그래픽
      TabIndex        =   16
      Top             =   1230
      Value           =   1  '확인
      Width           =   885
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "조회"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   15
      Top             =   1230
      Width           =   1365
   End
   Begin VB.TextBox txtSeqNo2 
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8760
      TabIndex        =   13
      Text            =   "1000"
      Top             =   300
      Width           =   1005
   End
   Begin VB.TextBox txtSeqNo1 
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7290
      TabIndex        =   12
      Text            =   "1"
      Top             =   300
      Width           =   1005
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1740
      TabIndex        =   8
      Top             =   300
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   23724033
      CurrentDate     =   38583
   End
   Begin VB.ComboBox cboOrder_Number 
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8340
      TabIndex        =   6
      Top             =   1170
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.ComboBox cboWorkStation 
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8340
      TabIndex        =   4
      Top             =   780
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.TextBox txtTemp 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   2
      Top             =   1170
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   3900
      Top             =   60
   End
   Begin VB.CommandButton cmdWorkList 
      Caption         =   "Order"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   1230
      Width           =   1365
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4890
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4380
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   5000
   End
   Begin FPSpread.vaSpread vasList1 
      Height          =   3375
      Left            =   270
      TabIndex        =   1
      Top             =   1800
      Width           =   9945
      _Version        =   196613
      _ExtentX        =   17542
      _ExtentY        =   5953
      _StockProps     =   64
      ColHeaderDisplay=   0
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
      MaxCols         =   8
      ScrollBars      =   2
      SpreadDesigner  =   "frmOrderInfo.frx":0442
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   3720
      TabIndex        =   9
      Top             =   300
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   23724033
      CurrentDate     =   38583
   End
   Begin FPSpread.vaSpread vasExam 
      Height          =   3855
      Left            =   6990
      TabIndex        =   25
      Top             =   5190
      Width           =   3225
      _Version        =   196613
      _ExtentX        =   5689
      _ExtentY        =   6800
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
      SpreadDesigner  =   "frmOrderInfo.frx":1DFB
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   3855
      Left            =   270
      TabIndex        =   26
      Top             =   5190
      Width           =   6705
      _Version        =   196613
      _ExtentX        =   11827
      _ExtentY        =   6800
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
      MaxCols         =   5
      ScrollBars      =   2
      SpreadDesigner  =   "frmOrderInfo.frx":6275
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "(1~60)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   195
      Left            =   2790
      TabIndex        =   24
      Top             =   870
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "초"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2490
      TabIndex        =   22
      Top             =   870
      Width           =   225
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "자동조회간격"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   870
      Width           =   1350
   End
   Begin VB.Image ImgOnOff 
      Height          =   525
      Left            =   5220
      Picture         =   "frmOrderInfo.frx":7AFA
      Top             =   630
      Width           =   525
   End
   Begin VB.Label lblOnOFF 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5850
      TabIndex        =   19
      Top             =   810
      Width           =   375
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8460
      TabIndex        =   14
      Top             =   360
      Width           =   120
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "작업번호"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5790
      TabIndex        =   11
      Top             =   360
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3540
      TabIndex        =   10
      Top             =   360
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "작업일자"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "차수"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6840
      TabIndex        =   5
      Top             =   1230
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "WorkStation"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6840
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   1320
   End
End
Attribute VB_Name = "frmOrderInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkMode_Click()
    If chkMode.Value = 1 Then
        chkMode.Caption = "Auto"
        SaveSetting "MEDIMATE", "LASCOrder", "SendMode", "1"
        Timer1.Enabled = True
    Else
        chkMode.Caption = "Manual"
        SaveSetting "MEDIMATE", "LASCOrder", "SendMode", "0"
        Timer1.Enabled = False
    End If
End Sub

Private Sub cmdClear_Click()
    ClearSpread vasList
    ClearSpread vasList1
    ClearSpread vasExam
End Sub

Private Sub cmdPortOpen_Click()
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
        cmdPortOpen.Caption = "PortOpen"
    Else
        cmdPortOpen.Caption = "PortClose"
        LASCPortOpen
    End If
End Sub

Sub LASCPortOpen()
    GetSetup_LASC
    
    MSComm1.CommPort = 1 'gSetup.Port
    MSComm1.Settings = gSetup.Speed & "," & gSetup.Parity & "," & gSetup.DataBit & "," & gSetup.StopBit
    If gSetup.DTREnable = "1" Then
        MSComm1.DTREnable = True
    Else
        MSComm1.DTREnable = False
    End If
    If gSetup.RTSEnable = "1" Then
        MSComm1.RTSEnable = True
    Else
        MSComm1.RTSEnable = False
    End If
    MSComm1.PortOpen = True
End Sub

Public Sub GetSetup_LASC()
    Dim db_tmp As String * 20
    Dim i As Integer
    Dim lRow As Long
       
    lRow = 0
    For i = 1 To 4
        db_tmp = ""
        Call GetPrivateProfileString("COM " & CStr(i), "Use", "", db_tmp, 20, App.Path & "\Interface.ini")
        txtTemp = Trim(db_tmp)
        If Trim(txtTemp) <> "" Then
'            lRow = lRow + 1
'
'            vasComList.Row = lRow
'            vasComList.Col = 1
'            If Trim(txtTemp) = "1" Then
'                vasComList.Value = 1
'            Else
'                vasComList.Value = 0
'            End If
            
            db_tmp = ""
            Call GetPrivateProfileString("COM " & CStr(i), "Gubun", "", db_tmp, 20, App.Path & "\Interface.ini")
            txtTemp = Trim(db_tmp)
            
            If Left(Trim(txtTemp), 4) = "LASC" Then
                
                gSetup.Port = i
                
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "Speed", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                gSetup.Speed = Trim(txtTemp)
            
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "Parity", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                gSetup.Parity = Trim(txtTemp)
            
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "DataBit", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                gSetup.DataBit = Trim(txtTemp)
            
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "StopBit", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                gSetup.StopBit = Trim(txtTemp)
            
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "RTSEnable", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                gSetup.RTSEnable = Trim(txtTemp)
            
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "DTREnable", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                gSetup.DTREnable = Trim(txtTemp)
            End If
        End If
    Next i

End Sub


Private Sub cmdSearch_Click()
    Dim lsDate As String
    Dim liOrdNo As Integer
    Dim i, j, k, n
    Dim lRow, lCol As Long
    
    Dim rsBarcode As ADODB.Recordset
    Dim cmdBarcode As New ADODB.Command
    
    Dim iOrd As Integer
    
    On Error GoTo errtrap
    
    k = DateDiff("d", DTPicker1.Value, DTPicker2)
    If k < 0 Then
        MsgBox "날짜 선택이 잘못되었습니다"
        Exit Sub
    End If
    
    Me.MousePointer = 11
    
    lsDate = DTPicker1.Value
    
    lRow = 1
    For i = 0 To k
        For j = 1 To 5
            With cmdSQL
                .ActiveConnection = cn_Ser
                .CommandType = adCmdStoredProc
                .CommandText = "Interface_WL_List_SELECT_sp"
                '.Parameters.Append .CreateParameter("retval", adInteger, adParamReturnValue)
                .Parameters.Append .CreateParameter("@i_instrumentcode", adChar, adParamInput, 11, "01")
                .Parameters.Append .CreateParameter("@i_WorkList_Date", adChar, adParamInput, 11, lsDate)
                .Parameters.Append .CreateParameter("@i_Order_Number", adChar, adParamInput, 11, j)
                .Parameters.Append .CreateParameter("@i_from_seq_number", adChar, adParamInput, 11, Trim(txtSeqNo1))
                .Parameters.Append .CreateParameter("@i_to_seq_number", adChar, adParamInput, 11, Trim(txtSeqNo2))
        
                Set rs = New ADODB.Recordset
                rs.CursorType = adOpenStatic
                Set rs = .Execute
            End With
        
            For n = 0 To cmdSQL.Parameters.Count - 1
                cmdSQL.Parameters.Delete 0
            Next n
            
            'lRow = 1
            While Not rs.EOF
                If vasList1.MaxRows < lRow Then
                    vasList1.MaxRows = lRow
                End If
                
                iOrd = -1
                With cmdSQL
                    .ActiveConnection = cn
                    .CommandType = adCmdText
                    .CommandText = "select barcode, OrdFlag from worklist where barcode = '" & Trim(rs.Fields.Item(1).Value) & "' "
                    Set rsBarcode = New ADODB.Recordset
                    Set rsBarcode = .Execute
                End With
                If Not rsBarcode.EOF Then
                    If Trim(CStr(rsBarcode.Fields.Item(1).Value)) = "B" Then
                        iOrd = 1
                    End If
                    rsBarcode.Close
                End If
                
                With cmdSQL
                    .ActiveConnection = cn
                    .CommandType = adCmdText
                    .CommandText = "select barcode from pat_res where barcode = '" & Trim(rs.Fields.Item(1).Value) & "' "
                    Set rsBarcode = New ADODB.Recordset
                    Set rsBarcode = .Execute
                End With
                If Not rsBarcode.EOF Then
                    If Trim(CStr(rsBarcode.Fields.Item(0).Value)) = Trim(rs.Fields.Item(1).Value) Then
                        iOrd = 1
                    End If
                    rsBarcode.Close
                End If
                
                If Not rsBarcode Is Nothing Then Set rsBarcode = Nothing
                If Not cmdBarcode Is Nothing Then Set cmdBarcode = Nothing
    
                If iOrd <> 1 Then
                    For lCol = 0 To rs.Fields.Count - 1
                        vasList1.Row = lRow
                        vasList1.Col = lCol + 2
                        If IsNull(rs.Fields.Item(lCol).Value) Then
                            vasList1.Text = ""
                        Else
                            vasList1.Text = Trim(CStr(rs.Fields.Item(lCol).Value))
                        End If
                    Next lCol
                    lRow = lRow + 1
                End If
                
                rs.MoveNext
            Wend
    
        Next j
        lsDate = Format(DateAdd("d", 1, CDate(lsDate)), "yyyy-mm-dd")
    Next i
    
    Me.MousePointer = 0
    
    Exit Sub
    
errtrap:
    Me.MousePointer = 0
    
    If Not rs Is Nothing Then Set rs = Nothing
    If Not cmdSQL Is Nothing Then Set cmdSQL = Nothing
    
    If Not rsBarcode Is Nothing Then Set rsBarcode = Nothing
    If Not cmdBarcode Is Nothing Then Set cmdBarcode = Nothing
    
    'MsgBox Err.Number & " : " & Err.Description

End Sub

Private Sub cmdSet_Click()
    Dim lsInt As String
    
    If Not IsNumeric(txtInt) Then
        MsgBox "1에서 60 사이의 숫자를 입력하십시오"
        txtInt.SetFocus
        Exit Sub
    End If
    If CInt(txtInt) < 1 Or CInt(txtInt) > 60 Then
        MsgBox "1에서 60 사이의 숫자를 입력하십시오"
        txtInt.SetFocus
        Exit Sub
    End If
    
    lsInt = Trim(txtInt)
    
    WritePrivateProfileString "Option", "Interval", lsInt, App.Path & "\Interface.ini"
    
    Timer1.Interval = CInt(txtInt) * 1000
'    If Timer1.Enabled = True Then
'        Timer1.Enabled = False
'        Timer1.Enabled = True
'    End If
End Sub

Private Sub cmdWorkList_Click()
Dim lsID As String
Dim i, j As Integer
Dim mExam As Variant
Dim AdoRs_Exam As ADODB.Recordset
Dim Ord(7) As String
Dim lsOrder As String
Dim lRow1, lRow As Long

If MSComm1.PortOpen = False Then
    LASCPortOpen
End If
    
If MSComm1.CTSHolding = False Then
    ImgOnOff.Picture = LoadPicture(App.Path & "\img\On.gif")
    Exit Sub
End If
    
For lRow1 = 1 To vasList1.DataRowCnt
    vasList1.Row = lRow1
    vasList1.Col = 1
    If vasList1.Value = 0 Then
        
        For i = 1 To 7
            Ord(i) = "0"
        Next i
        
        lRow = vasList.DataRowCnt + 1
        If lRow > vasList.MaxRows Then
            vasList.MaxRows = lRow
        End If
        
        lsID = Trim(GetText(vasList1, lRow1, 3))
        
        SetText vasList, lsID, lRow, 1
        SetText vasList, "A", lRow, 2
        SetText vasList, Trim(GetText(vasList1, lRow1, 4)), lRow, 3
        SetText vasList, Trim(GetText(vasList1, lRow1, 5)), lRow, 4
        SetText vasList, GetDateFull, lRow, 5
                
        mExam = Get_OrderBody1(lsID)
        If Not IsNull(mExam) Then
            SQL = "select equipcode, examcode, examname, OrdGubun from equipexam where equip = '" & gEquip & "' and OrdGubun <> 'N' "
            Set AdoRs_Exam = db_select_rs(gLocal, SQL)
            
            ClearSpread vasExam
            For j = LBound(mExam, 2) To UBound(mExam, 2)
                SetText vasExam, mExam(3, j), j, 1
                SetText vasExam, mExam(4, j), j, 2
                
                If Not AdoRs_Exam Is Nothing Then
                    AdoRs_Exam.MoveFirst
                    Do Until AdoRs_Exam.EOF
                        If Trim(AdoRs_Exam("examcode")) = mExam(3, j) Then
                            Select Case Trim(AdoRs_Exam("OrdGubun"))
                            Case "C": Ord(1) = "1"
                            Case "D": Ord(2) = "1"
                            Case "R": Ord(3) = "1"
                            Case "P": Ord(4) = "1"
                            Case "S": Ord(5) = "1"
                            Case "X": Ord(6) = "1"
                            Case "B": Ord(7) = "1"
                            End Select
                            
                            Exit Do
                        End If
                        
                        AdoRs_Exam.MoveNext
                    Loop
                End If
            Next j
            
            lsOrder = ""
            For i = 1 To 7
                lsOrder = lsOrder & Ord(i)
            Next i
            
            'MsgBox lsOrder
            
            If lsOrder <> "0000000" And lsOrder <> "" Then
                lsOrder = chrSTX & "S" & "00000000" & SetChar(lsID, 15, 1, " ") & "********" & lsOrder & "00000"
                lsOrder = lsOrder & "0000000000000"
                lsOrder = lsOrder & "0000000000000"
                lsOrder = lsOrder & "0000000000000"
                lsOrder = lsOrder & "000****************************************" & chrETX
                MSComm1.Output = lsOrder
                
                
                SQL = "Insert Into WorkList(ReceDate, Barcode, PID, PName, OrdFlag, OrdDateTime, ResDateTime, OrdCnt ) " & vbCrLf & _
                      "Values ('" & Trim(GetText(vasList, lRow, 5)) & "', '" & lsID & "', '" & Trim(GetText(vasList, lRow, 3)) & "','" & Trim(GetText(vasList, lRow, 4)) & "', 'B','','', 0) "
                res = SendQuery(gLocal, SQL)
                If res = 1 Then
                    SetText vasList, "B", lRow, 2
                    DeleteRow vasList1, lRow1, lRow1
                Else
                    SaveQuery SQL
                    'Exit Function
                End If
            End If
            
        End If
    
    End If
Next lRow1

MSComm1.PortOpen = False

End Sub


Function GetSetup() As Boolean
'---------------------------------------------------------------------------------------------------------------------
'                       Setup  File을 읽어온다.
'---------------------------------------------------------------------------------------------------------------------
    Dim db_tmp As String * 20

    db_tmp = ""
    
    GetSetup = False
    
    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "driver", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.Driver = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "uid", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.User = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "pwd", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.Passwd = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "server", "", db_tmp, 100, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.Server = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "database", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.DB = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "hostname", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.HostName = Trim(txtTemp)

    GetSetup = True

End Function
    
Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    Dim db_tmp As String * 100
    Dim lsData As String
    Dim i As Integer
    
    db_tmp = ""
    Call GetPrivateProfileString("Option", "WorkStationCode", "", db_tmp, 100, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    lsData = Trim(txtTemp)
    
    cboWorkStation.Clear
    i = InStr(1, lsData, ",")
    Do While i > 0
        cboWorkStation.AddItem Trim(Left(lsData, i - 1))
        lsData = Mid(lsData, i + 1)
        i = InStr(1, lsData, ",")
    Loop
    If Trim(lsData) <> "" Then
        cboWorkStation.AddItem Trim(lsData)
    End If
    
    db_tmp = ""
    Call GetPrivateProfileString("Option", "WorkStationCode", "", db_tmp, 100, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    lsData = Trim(txtTemp)
    
    cboOrder_Number.Clear
    i = InStr(1, lsData, ",")
    Do While i > 0
        cboOrder_Number.AddItem Trim(Left(lsData, i - 1))
        lsData = Mid(lsData, i + 1)
        i = InStr(1, lsData, ",")
    Loop
    If Trim(lsData) <> "" Then
        cboOrder_Number.AddItem Trim(lsData)
    End If
    
    db_tmp = ""
    Call GetPrivateProfileString("Option", "Interval", "", db_tmp, 100, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    txtInt = Trim(txtTemp)
    
    
    cn_Local_Flag = False
    cn_Server_Flag = False
    
    GetSetup
    
    If Connect_Local Then
        cn_Local_Flag = True
    End If
    
    If Connect_Server Then
        cn_Server_Flag = True
    End If
    
    cmdPortOpen_Click
    
    If Trim(GetSetting("MEDIMATE", "LASCOrder", "SendMode", "0")) = "1" Then
        chkMode.Value = 1
        Timer1.Enabled = True
    Else
        chkMode.Value = 0
        Timer1.Enabled = False
    End If
    
    DTPicker1.Value = Format(CDate(GetDateFull), "yyyy-mm-dd")
    DTPicker2.Value = DTPicker1.Value
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim lsInt As String
    lsInt = Trim(txtInt)
    WritePrivateProfileString "Option", "Interval", lsInt, App.Path & "\Interface.ini"

    Timer1.Enabled = False
        
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    
    DisConnect_Server
    
    DisConnect_Local

End Sub

Private Sub Timer1_Timer()
    If MSComm1.CTSHolding = True Then
        ImgOnOff.Picture = LoadPicture(App.Path & "\IMG\On.gif")
        
        cmdSearch_Click
        cmdWorkList_Click
    Else
        ImgOnOff.Picture = LoadPicture(App.Path & "\IMG\Off.gif")
    End If
End Sub
