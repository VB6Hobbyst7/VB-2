VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmRemark 
   Caption         =   "소견등록"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4695
   ScaleWidth      =   9585
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame Frame1 
      Height          =   4395
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   9405
      Begin VB.CheckBox chkCmnt 
         Caption         =   "소견표시"
         Height          =   405
         Left            =   4170
         TabIndex        =   12
         Top             =   240
         Width           =   1425
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6420
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtCmntP3 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Left            =   210
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   6600
         Visible         =   0   'False
         Width           =   8955
      End
      Begin VB.TextBox txtCmntP2 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Left            =   210
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   4650
         Visible         =   0   'False
         Width           =   8955
      End
      Begin VB.TextBox txtCmntP1 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Left            =   210
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   2700
         Width           =   8955
      End
      Begin VB.TextBox txtCmntN 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Left            =   210
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   750
         Width           =   8955
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7830
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "[Total IgE 증가 : Class 2 이상소견]"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   270
         TabIndex        =   8
         Top             =   6330
         Visible         =   0   'False
         Width           =   3585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "[Total IgE 증가 : 음성소견]"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   270
         TabIndex        =   6
         Top             =   4410
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "[10개 이상 항원 코멘트]"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   270
         TabIndex        =   3
         Top             =   2460
         Width           =   2295
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "[Total IgE 증가코멘트]"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   270
         TabIndex        =   2
         Top             =   480
         Width           =   2235
      End
   End
   Begin FPSpread.vaSpread vasTemp 
      Height          =   975
      Left            =   11490
      TabIndex        =   10
      Top             =   1500
      Visible         =   0   'False
      Width           =   1695
      _Version        =   393216
      _ExtentX        =   2990
      _ExtentY        =   1720
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
      SpreadDesigner  =   "frmRemark.frx":0000
   End
End
Attribute VB_Name = "frmRemark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim i As Integer
    
    
    If chkCmnt.Value = "1" Then
        Call WritePrivateProfileString("ASSAY", "CMTVIEW", "1", App.Path & "\Interface.ini")
    Else
        Call WritePrivateProfileString("ASSAY", "CMTVIEW", "0", App.Path & "\Interface.ini")
    End If
    
    SQL = "DELETE FROM CONFIG WHERE CATEGORY = 'COMMENT'"
          
    Res = SendQuery(gLocal, SQL)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If

    For i = 1 To 4
        Select Case i
            Case 1:
                      SQL = "INSERT INTO CONFIG (CATEGORY,TITLE,CONTENT,DESCRIPTION) VALUES " & vbCr
                SQL = SQL & " ('COMMENT','N','" & txtCmntN & "','" & Format(Now, "yyyymmdd") & "')"
            Case 2:
                      SQL = "INSERT INTO CONFIG (CATEGORY,TITLE,CONTENT,DESCRIPTION) VALUES " & vbCr
                SQL = SQL & " ('COMMENT','P1','" & txtCmntP1 & "','" & Format(Now, "yyyymmdd") & "')"
            Case 3:
                      SQL = "INSERT INTO CONFIG (CATEGORY,TITLE,CONTENT,DESCRIPTION) VALUES " & vbCr
                SQL = SQL & " ('COMMENT','P2','" & txtCmntP2 & "','" & Format(Now, "yyyymmdd") & "')"
            Case 4:
                      SQL = "INSERT INTO CONFIG (CATEGORY,TITLE,CONTENT,DESCRIPTION) VALUES " & vbCr
                SQL = SQL & " ('COMMENT','P3','" & txtCmntP3 & "','" & Format(Now, "yyyymmdd") & "')"
        End Select
        Res = SendQuery(gLocal, SQL)
        
        If Res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If
    Next

End Sub

Private Sub Form_Load()
    ClearText
    DisplayList
End Sub

Private Sub ClearText()

    txtCmntN = ""
    txtCmntP1 = ""
    txtCmntP2 = ""
    txtCmntP3 = ""
    
'    txtCmntN = CMNT.N
'    txtCmntP1 = CMNT.P1
'    txtCmntP2 = CMNT.P2
'    txtCmntP3 = CMNT.P3
    
End Sub

Private Sub DisplayList()
    Dim i As Integer
    
    chkCmnt.Value = gAssayNM.CMTVIEW

          SQL = " Select TITLE, CONTENT " & vbCr
    SQL = SQL & "  FROM CONFIG  " & vbCr
    SQL = SQL & " Where CATEGORY = 'COMMENT'"
          
    Res = GetDBSelectVas(gLocal, SQL, vasTemp)
        
    If Res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
            
    vasTemp.MaxRows = vasTemp.DataRowCnt + 1

    '-- 서버로 결과값 저장하기
    For i = 1 To vasTemp.DataRowCnt
        Select Case Trim(GetText(vasTemp, i, 1))
            Case "N": txtCmntN = Trim(GetText(vasTemp, i, 2))
            Case "P1": txtCmntP1 = Trim(GetText(vasTemp, i, 2))
            Case "P2": txtCmntP2 = Trim(GetText(vasTemp, i, 2))
            Case "P3": txtCmntP3 = Trim(GetText(vasTemp, i, 2))
        End Select
    Next
    
End Sub

