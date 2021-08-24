VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form FrmViewExamInfo 
   Caption         =   "의료정보사항"
   ClientHeight    =   5310
   ClientLeft      =   1020
   ClientTop       =   3480
   ClientWidth     =   10740
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   10740
   Begin FPSpreadADO.fpSpread SS1 
      Height          =   4860
      Left            =   45
      TabIndex        =   2
      Top             =   450
      Width           =   5295
      _Version        =   196608
      _ExtentX        =   9340
      _ExtentY        =   8573
      _StockProps     =   64
      AutoSize        =   -1  'True
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      EditModePermanent=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      MaxCols         =   2
      MaxRows         =   20
      MoveActiveOnFocus=   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "FrmViewExamInfo.frx":0000
      UserResize      =   0
      Appearance      =   2
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   330
      Left            =   5400
      TabIndex        =   1
      Top             =   45
      Width           =   5325
      _Version        =   65536
      _ExtentX        =   9393
      _ExtentY        =   582
      _StockProps     =   15
      Caption         =   "보험법 인정기준"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   330
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   5325
      _Version        =   65536
      _ExtentX        =   9393
      _ExtentY        =   582
      _StockProps     =   15
      Caption         =   "진료정보"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      RoundedCorners  =   0   'False
   End
   Begin FPSpreadADO.fpSpread SS2 
      Height          =   4860
      Left            =   5400
      TabIndex        =   3
      Top             =   450
      Width           =   5295
      _Version        =   196608
      _ExtentX        =   9340
      _ExtentY        =   8573
      _StockProps     =   64
      AutoSize        =   -1  'True
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      EditModePermanent=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      MaxCols         =   2
      MaxRows         =   20
      MoveActiveOnFocus=   0   'False
      Protect         =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "FrmViewExamInfo.frx":037C
      UserResize      =   0
      Appearance      =   2
   End
End
Attribute VB_Name = "FrmViewExamInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strRemarks(40)              As String * 50


Private Sub Form_Activate()
    Dim i                       As Integer
    Dim j                       As Integer
    Dim nRow                    As Integer
    Dim nRowCnt                 As Integer
    Dim strRemark               As String
    Dim nSS1Row                 As Integer
    Dim strTitleName            As String
    
    Me.Refresh
    
    strSql = ""
    strSql = strSql & " SELECT * "
    strSql = strSql & " FROM TW_MIS_OCS.TWOCS_OINFOR "
    strSql = strSql & " WHERE OrderCode = '" & GstrSELECTOrderCode & "'"
    
    If adoSetOpen(strSql, adoSet) Then
        Do Until adoSet.EOF
            Select Case Trim(adoSet.Fields("GbData").Value & "")
                Case "1":  GoSub Display_Bohum   '보험법 인정기준
                Case Else: GoSub Display_Infor   '약,검사정보
            End Select
            adoSet.MoveNext
        Loop
        Call adoSetClose(adoSet)
    End If
    
    
    Exit Sub
    
    
'/-----------------------------------------------------------------------------------/

Display_Bohum:

    strRemark = adoSet.Fields("Remark1").Value & ""
    If Trim(strRemark) <> "" Then
        GoSub Find_Row_Count
        If nRowCnt > SS2.MaxRows Then SS2.MaxRows = nRowCnt
        For j = 1 To nRowCnt
            SS2.Row = j
            SS2.Col = 1:        SS2.Text = strRemarks(j)
        Next j
    End If
    
    Return
    

'/-----------------------------------------------------------------------------------/

Display_Infor:
    
    
    nRow = 1:                   nSS1Row = 1
    
    
    strTitleName = adoSet.Fields("TitleName1").Value & ""
    strRemark = Trim(adoSet.Fields("Remark1").Value & "")
    If Trim(strRemark) <> "" Then
        GoSub Display_Title
        GoSub Find_Row_Count
        GoSub Display_Rows
    End If
    
    strTitleName = adoSet.Fields("TitleName2").Value & ""
    strRemark = Trim(adoSet.Fields("Remark2").Value & "")
    If Trim(strRemark) <> "" Then
        GoSub Display_Title
        GoSub Find_Row_Count
        GoSub Display_Rows
    End If
    
    strTitleName = adoSet.Fields("TitleName3").Value & ""
    strRemark = Trim(adoSet.Fields("Remark3").Value & "")
    If Trim(strRemark) <> "" Then
        GoSub Display_Title
        GoSub Find_Row_Count
        GoSub Display_Rows
    End If
    
    strTitleName = adoSet.Fields("TitleName4").Value & ""
    strRemark = Trim(adoSet.Fields("Remark4").Value & "")
    If Trim(strRemark) <> "" Then
        GoSub Display_Title
        GoSub Find_Row_Count
        GoSub Display_Rows
    End If
        
    strTitleName = adoSet.Fields("TitleName5").Value & ""
    strRemark = Trim(adoSet.Fields("Remark5").Value & "")
    If Trim(strRemark) <> "" Then
        GoSub Display_Title
        GoSub Find_Row_Count
        GoSub Display_Rows
    End If

    Return


'/-----------------------------------------------------------------------------------/

Display_Title:
        
    SS1.Row = nRow
    SS1.Col = 1:            SS1.Text = Format(nSS1Row, "0") & "."
    SS1.Col = 2:            SS1.Text = strTitleName
    SS1.Row = nRow:         SS1.Row2 = nRow
    SS1.Col = 1:            SS1.Col2 = SS1.MaxCols
    SS1.BlockMode = True
    SS1.ForeColor = RGB(0, 0, 255)
    SS1.BlockMode = False
    nSS1Row = nSS1Row + 1

    Return


'/-----------------------------------------------------------------------------------/

Display_Rows:
    
    If nRowCnt + nRow > SS1.MaxRows Then SS1.MaxRows = nRowCnt + nRow
    
    For j = 1 To nRowCnt
        SS1.Row = nRow + j
        SS1.Col = 2:    SS1.Text = Trim(strRemarks(j))
    Next j
    
    nRow = nRow + nRowCnt + 1
    
    Return
    
    
'/-----------------------------------------------------------------------------------/

Find_Row_Count:

    Dim nLen                    As Integer
    Dim nLenMax                 As Integer
    Dim nInx                    As Integer
    
    nRowCnt = 0
    nLenMax = LenB(Trim(strRemark))
    
    Do
        nRowCnt = nRowCnt + 1
        strRemarks(nRowCnt) = strRemark
        nLen = InStrB(1, strRemark, vbCrLf)
        If nLen <> 0 Then
            strRemarks(nRowCnt) = LeftB$(strRemark, nLen - 1)
            strRemark = Trim(MidB$(strRemark, nLen + 4))
        End If
    
    Loop Until nLenMax <= nLen + 2 Or nLen = 0
        
    Return

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Set FrmViewExamInfo = Nothing
    Unload Me

End Sub
