VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmFunc 
   BorderStyle     =   3  '≈©±‚ ∞Ì¡§ ¥Î»≠ ªÛ¿⁄
   Caption         =   "≈∞ ∏  ∞¸∏Æ"
   ClientHeight    =   3900
   ClientLeft      =   1815
   ClientTop       =   2370
   ClientWidth     =   7275
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3900
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2580
      Left            =   135
      TabIndex        =   2
      Top             =   630
      Width           =   6900
      _Version        =   65536
      _ExtentX        =   12171
      _ExtentY        =   4551
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±º∏≤"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Begin VB.TextBox txtMapString 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   1
         Left            =   735
         MaxLength       =   20
         TabIndex        =   14
         Top             =   210
         Width           =   2580
      End
      Begin VB.TextBox txtMapString 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   7
         Left            =   4155
         MaxLength       =   20
         TabIndex        =   13
         Top             =   210
         Width           =   2580
      End
      Begin VB.TextBox txtMapString 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   2
         Left            =   735
         MaxLength       =   20
         TabIndex        =   12
         Top             =   585
         Width           =   2580
      End
      Begin VB.TextBox txtMapString 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   8
         Left            =   4155
         MaxLength       =   20
         TabIndex        =   11
         Top             =   585
         Width           =   2580
      End
      Begin VB.TextBox txtMapString 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   3
         Left            =   735
         MaxLength       =   20
         TabIndex        =   10
         Top             =   945
         Width           =   2580
      End
      Begin VB.TextBox txtMapString 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   9
         Left            =   4155
         MaxLength       =   20
         TabIndex        =   9
         Top             =   945
         Width           =   2580
      End
      Begin VB.TextBox txtMapString 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   4
         Left            =   735
         MaxLength       =   20
         TabIndex        =   8
         Top             =   1320
         Width           =   2580
      End
      Begin VB.TextBox txtMapString 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   10
         Left            =   4155
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1320
         Width           =   2580
      End
      Begin VB.TextBox txtMapString 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   5
         Left            =   735
         MaxLength       =   20
         TabIndex        =   6
         Top             =   1680
         Width           =   2580
      End
      Begin VB.TextBox txtMapString 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   11
         Left            =   4140
         MaxLength       =   20
         TabIndex        =   5
         Top             =   1680
         Width           =   2580
      End
      Begin VB.TextBox txtMapString 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   6
         Left            =   735
         MaxLength       =   20
         TabIndex        =   4
         Top             =   2055
         Width           =   2580
      End
      Begin VB.TextBox txtMapString 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   12
         Left            =   4155
         MaxLength       =   20
         TabIndex        =   3
         Top             =   2055
         Width           =   2580
      End
      Begin Threed.SSPanel panKey 
         Height          =   360
         Index           =   0
         Left            =   135
         TabIndex        =   15
         Top             =   180
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   635
         _StockProps     =   15
         Caption         =   "F1"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   5
         BorderWidth     =   1
         Outline         =   -1  'True
         Font3D          =   3
      End
      Begin Threed.SSPanel panKey 
         Height          =   360
         Index           =   1
         Left            =   3555
         TabIndex        =   16
         Top             =   180
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   635
         _StockProps     =   15
         Caption         =   "F7"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   5
         BorderWidth     =   1
         Outline         =   -1  'True
         Font3D          =   3
      End
      Begin Threed.SSPanel panKey 
         Height          =   360
         Index           =   2
         Left            =   135
         TabIndex        =   17
         Top             =   555
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   635
         _StockProps     =   15
         Caption         =   "F2"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   5
         BorderWidth     =   1
         Outline         =   -1  'True
         Font3D          =   3
      End
      Begin Threed.SSPanel panKey 
         Height          =   360
         Index           =   3
         Left            =   3555
         TabIndex        =   18
         Top             =   555
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   635
         _StockProps     =   15
         Caption         =   "F8"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   5
         BorderWidth     =   1
         Outline         =   -1  'True
         Font3D          =   3
      End
      Begin Threed.SSPanel panKey 
         Height          =   360
         Index           =   4
         Left            =   135
         TabIndex        =   19
         Top             =   930
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   635
         _StockProps     =   15
         Caption         =   "F3"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   5
         BorderWidth     =   1
         Outline         =   -1  'True
         Font3D          =   3
      End
      Begin Threed.SSPanel panKey 
         Height          =   360
         Index           =   5
         Left            =   3555
         TabIndex        =   20
         Top             =   930
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   635
         _StockProps     =   15
         Caption         =   "F9"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   5
         BorderWidth     =   1
         Outline         =   -1  'True
         Font3D          =   3
      End
      Begin Threed.SSPanel panKey 
         Height          =   360
         Index           =   6
         Left            =   135
         TabIndex        =   21
         Top             =   1305
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   635
         _StockProps     =   15
         Caption         =   "F4"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   5
         BorderWidth     =   1
         Outline         =   -1  'True
         Font3D          =   3
      End
      Begin Threed.SSPanel panKey 
         Height          =   360
         Index           =   7
         Left            =   3555
         TabIndex        =   22
         Top             =   1305
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   635
         _StockProps     =   15
         Caption         =   "F10"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   5
         BorderWidth     =   1
         Outline         =   -1  'True
         Font3D          =   3
      End
      Begin Threed.SSPanel panKey 
         Height          =   360
         Index           =   8
         Left            =   135
         TabIndex        =   23
         Top             =   1680
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   635
         _StockProps     =   15
         Caption         =   "F5"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   5
         BorderWidth     =   1
         Outline         =   -1  'True
         Font3D          =   3
      End
      Begin Threed.SSPanel panKey 
         Height          =   360
         Index           =   9
         Left            =   3555
         TabIndex        =   24
         Top             =   1680
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   635
         _StockProps     =   15
         Caption         =   "F11"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   5
         BorderWidth     =   1
         Outline         =   -1  'True
         Font3D          =   3
      End
      Begin Threed.SSPanel panKey 
         Height          =   360
         Index           =   10
         Left            =   135
         TabIndex        =   25
         Top             =   2055
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   635
         _StockProps     =   15
         Caption         =   "F6"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   5
         BorderWidth     =   1
         Outline         =   -1  'True
         Font3D          =   3
      End
      Begin Threed.SSPanel panKey 
         Height          =   360
         Index           =   11
         Left            =   3555
         TabIndex        =   26
         Top             =   2055
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   635
         _StockProps     =   15
         Caption         =   "F12"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   5
         BorderWidth     =   1
         Outline         =   -1  'True
         Font3D          =   3
      End
   End
   Begin MSForms.CommandButton cmdExit 
      Height          =   510
      Left            =   5580
      TabIndex        =   28
      Top             =   3285
      Width           =   1455
      Caption         =   "¡æ∑·"
      PicturePosition =   327683
      Size            =   "2566;900"
      Picture         =   "frmFunc.frx":0000
      FontName        =   "±º∏≤"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdAdd 
      Height          =   510
      Left            =   4095
      TabIndex        =   27
      Top             =   3285
      Width           =   1500
      Caption         =   "¿˙¿Â"
      PicturePosition =   327683
      Size            =   "2646;900"
      Picture         =   "frmFunc.frx":031A
      FontName        =   "±º∏≤"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Label LblExamName 
      Alignment       =   2  '∞°øÓµ• ∏¬√„
      Caption         =   "¿”ªÛ»≠«– ∞ÀªÁ 1"
      BeginProperty Font 
         Name            =   "±º∏≤√º"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   225
      Left            =   3165
      TabIndex        =   1
      Top             =   300
      Width           =   1830
   End
   Begin VB.Label Label1 
      Alignment       =   2  '∞°øÓµ• ∏¬√„
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  '¥‹¿œ ∞Ì¡§
      Caption         =   "Result Key Mapping"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   2985
   End
End
Attribute VB_Name = "frmFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()

    Dim ii              As Integer
    Dim cCodeGu         As String
    Dim cCodeky         As String
    Dim cCodeNm         As String
    Dim sFSLip          As String
    
    sFSLip = Left(frmResult.cmbSLip.Text, 2)
    
    
    For ii = 1 To 12
        cCodeGu = "19"
        cCodeky = sFSLip & (ii + 111)
        cCodeNm = txtMapString(ii).Text
        
        gStrSql = ""
        gStrSql = gStrSql & " SELECT * "
        gStrSql = gStrSql & " FROM   TWEXAM_SPECODE    "
        gStrSql = gStrSql & " WHERE  CodeGu      =    '19'    "
        gStrSql = gStrSql & " AND    CodeKy      =    '" & sFSLip & (ii + 111) & "' "
        
        If False = adoSetOpen(gStrSql, adoSet) And txtMapString(ii).Text <> "" Then
            gStrSql = ""
            gStrSql = gStrSql & " INSERT INTO TWEXAM_SPECODE"
            gStrSql = gStrSql & "        ( CodeGu, CodeKy, CodeNm )"
            gStrSql = gStrSql & " VALUES ( '" & cCodeGu & "',"
            gStrSql = gStrSql & "          '" & cCodeky & "',"
            gStrSql = gStrSql & "          '" & cCodeNm & "')"
            adoConnect.BeginTrans
            If adoExec(gStrSql) Then
                adoConnect.CommitTrans
            Else
                adoConnect.RollbackTrans
            End If
        Else
            Call adoSetClose(adoSet)
            gStrSql = ""
            gStrSql = gStrSql & " UPDATE TWEXAM_SPECODE    "
            gStrSql = gStrSql & " SET    CodeNm      =    '" & cCodeNm & "'"
            gStrSql = gStrSql & " WHERE  CodeGu      =    '19'"
            gStrSql = gStrSql & " AND    CodeKy      =    '" & sFSLip & (ii + 111) & "'"
            adoConnect.BeginTrans
            If adoExec(gStrSql) Then
                adoConnect.CommitTrans
            Else
                adoConnect.RollbackTrans
            End If
        End If
    Next ii
   

End Sub


Private Sub cmdExit_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

    Dim ii              As Integer
    Dim iSLip           As String
    
 
    LblExamName.Caption = GsExamJong
    iSLip = Left(frmResult.cmbSLip.Text, 2)
    
    
    gStrSql = ""
    gStrSql = gStrSql & " SELECT CodeKy,CodeNm FROM TWEXAM_SPECODE "
    gStrSql = gStrSql & " WHERE  CodeGu = '19' "
    gStrSql = gStrSql & " AND    CodeKy Like '" & iSLip & "%'  "
    gStrSql = gStrSql & " ORDER  BY CodeKy  "
    
    If False = adoSetOpen(gStrSql, adoSet) Then Exit Sub
    
    For ii = 0 To adoSet.RecordCount - 1
        txtMapString(Val(Mid$(adoSet.Fields("CodeKy").Value & "", 3, 3)) - 111).Text = adoSet.Fields("CodeNm").Value & ""
        adoSet.MoveNext
    Next ii
    Call adoSetClose(adoSet)
    
End Sub

Private Sub txtMapString_GotFocus(Index As Integer)
   
   txtMapString(Index).SelStart = 0
   txtMapString(Index).SelLength = Len(txtMapString(Index).Text)

End Sub


Private Sub txtMapString_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

    KeyAscii = 0
 
    SendKeys "{tab}"

End Sub


