VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form Frm_New 
   BorderStyle     =   1  '단일 고정
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15345
   Icon            =   "Frm_New.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   15345
   Begin VB.Frame Fam_B 
      Caption         =   "바코드 - 포맷 생성"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7365
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   14400
      Begin VB.ComboBox Cbo_HeadDarkness 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Frm_New.frx":628A
         Left            =   11895
         List            =   "Frm_New.frx":62E8
         TabIndex        =   18
         Text            =   "15"
         Top             =   2940
         Width           =   1710
      End
      Begin VB.ComboBox Cbo_PrinterSpeed 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Frm_New.frx":6365
         Left            =   11895
         List            =   "Frm_New.frx":637B
         TabIndex        =   16
         Text            =   " 3"
         Top             =   2415
         Width           =   1710
      End
      Begin VB.ComboBox cbo_Port 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Frm_New.frx":6397
         Left            =   11895
         List            =   "Frm_New.frx":63B6
         TabIndex        =   10
         Text            =   "COM1"
         Top             =   525
         Width           =   1710
      End
      Begin VB.ComboBox Cbo_Baud 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Frm_New.frx":63F0
         Left            =   11895
         List            =   "Frm_New.frx":6403
         TabIndex        =   9
         Text            =   "9600"
         Top             =   1245
         Width           =   1710
      End
      Begin VB.ComboBox Cbo_Dpi 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   11895
         TabIndex        =   8
         Text            =   "200 dpi"
         Top             =   1920
         Width           =   1710
      End
      Begin VB.TextBox Txt_CenterX 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11895
         TabIndex        =   7
         Text            =   "0"
         Top             =   3570
         Width           =   810
      End
      Begin VB.TextBox Txt_CenterY 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   11895
         TabIndex        =   6
         Text            =   "0"
         Top             =   3915
         Width           =   810
      End
      Begin FPSpread.vaSpread Spr_B 
         Height          =   6825
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   9885
         _Version        =   393216
         _ExtentX        =   17436
         _ExtentY        =   12039
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   6
         MaxRows         =   15
         RowHeaderDisplay=   0
         ScrollBarMaxAlign=   0   'False
         ScrollBars      =   0
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   3
         ShadowColor     =   16761024
         SpreadDesigner  =   "Frm_New.frx":642A
      End
      Begin VB.Label Label5 
         Caption         =   "해드 온도"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10800
         TabIndex        =   19
         Top             =   2970
         Width           =   915
      End
      Begin VB.Label Label4 
         Caption         =   "프린터 스피드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10380
         TabIndex        =   17
         Top             =   2445
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "통신 포트 설정"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10260
         TabIndex        =   15
         Top             =   525
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "데이터 전송 속도"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10080
         TabIndex        =   14
         Top             =   1260
         Width           =   1680
      End
      Begin VB.Label Label3 
         Caption         =   "장비 DPI 값"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10575
         TabIndex        =   13
         Top             =   1980
         Width           =   1140
      End
      Begin VB.Label Label7 
         Caption         =   "원점 - X(mm)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10395
         TabIndex        =   12
         Top             =   3600
         Width           =   1365
      End
      Begin VB.Label Label8 
         Caption         =   "원점 - Y"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10305
         TabIndex        =   11
         Top             =   3990
         Width           =   990
      End
   End
   Begin VB.Frame Fam_C 
      Caption         =   "바코드 - 레코드 처리"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   90
      TabIndex        =   2
      Top             =   900
      Visible         =   0   'False
      Width           =   15165
      Begin VB.CommandButton Cmd_Printer 
         Caption         =   "발행"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   11820
         TabIndex        =   5
         Top             =   6690
         Width           =   3090
      End
      Begin FPSpread.vaSpread Spr_C 
         Height          =   6315
         Left            =   60
         TabIndex        =   4
         Top             =   270
         Width           =   15030
         _Version        =   393216
         _ExtentX        =   26511
         _ExtentY        =   11139
         _StockProps     =   64
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   14
         MaxRows         =   30
         ShadowColor     =   16761024
         SpreadDesigner  =   "Frm_New.frx":11338
      End
   End
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   7890
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   13917
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "바코드 - 포맷 생성"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "바코드 - 레코드 처리"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Frm_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
'***  Description   : Printer 발행
'***  Modification Log : 2007/06/14  김동호  Initial Coding
'***********************************************************************************
Private Sub Cmd_Printer_Click()

    Cmd_Printer.Enabled = False

    Dim i, j, k As Integer
    Dim Ls_SprB(4) As String
    Dim Ls_SprC(12) As String
    Dim Ls_KSP As String
    Dim Pri_Sw As Integer '// 증정품 Check
    Dim Gift_Pr1, Gift_Pr2, Gift_Pr3 As String
  ' Frm_Main.Mcom.PortOpen = True
   For i = 1 To Me.Spr_C.MaxRows Step 1
        Me.Spr_C.Row = i
        Me.Spr_C.Col = 1
        
        If Trim(Me.Spr_C.Text) = "1" Then
            For k = 2 To Me.Spr_C.MaxCols Step 1
                Me.Spr_C.Col = k
                Ls_SprC(k - 2) = Me.Spr_C.Text
            Next k
        
            Call Zebra_sTr
        
            For j = 1 To Me.Spr_B.MaxRows Step 1
                Me.Spr_B.Row = j
                Me.Spr_B.Col = 1
                
                If Trim(Me.Spr_B.Text) = "1" Then
                    For k = 2 To Me.Spr_B.MaxCols Step 1
                        Me.Spr_B.Col = k
                        Ls_SprB(k - 2) = Me.Spr_B.Text
                    Next k
                
                    Select Case j
                        Case 1  '// 바코드상
                            Call Zebra_BarN(Ls_SprB(0), Ls_SprB(1), Ls_SprB(2), Ls_SprB(3), Ls_SprC(0) & Ls_SprC(1) & Ls_SprC(2) & Ls_SprC(3))
                        Case 2  '// C/N0
                            Call Zebra_A0(Ls_SprB(0), Ls_SprB(1), Ls_SprB(2), Ls_SprB(3), Ls_SprC(0) & "-" & Ls_SprC(1) & "-" & Ls_SprC(2) & "-" & Ls_SprC(3))
                        Case 3  '// Size 영문인쇄
                            Call Zebra_A0(Val(Ls_SprB(0)), Val(Ls_SprB(1)), Val(Ls_SprB(2)), Val(Ls_SprB(3)), Ls_SprC(4))
'                            Call Label_Print(Ls_SprC(4), Val(Ls_SprB(0)), Val(Ls_SprB(1)), Val(Ls_SprB(2)))
                        Case 4  '// 색상
                            If Trim(Color_Shr(Ls_SprC(2))) = "" Then
                                MsgBox "Color 코드가 존재 하지 않습니다.", vbCritical, "Code Error"
                                Frm_Main.Mcom.PortOpen = False
                                Cmd_Printer.Enabled = True
                                Exit Sub
                            Else
'                                Call Zebra_Han(Val(Ls_SprB(0)), Val(Ls_SprB(1)), Val(Ls_SprB(2)), Val(Ls_SprB(3)), Color_Shr(Ls_SprC(2)))
                                Call Label_Print(Color_Shr(Ls_SprC(2)), Val(Ls_SprB(0)), Val(Ls_SprB(1)), Val(Ls_SprB(2)))
                            End If
                        Case 5  '// 품명
'                            Call Zebra_Han(Val(Ls_SprB(0)), Val(Ls_SprB(1)), Val(Ls_SprB(2)), Val(Ls_SprB(3)), Ls_SprC(5))
                            Call Label_Print(Ls_SprC(5), Val(Ls_SprB(0)), Val(Ls_SprB(1)), Val(Ls_SprB(2)))
                        Case 6  '// 거리
                            Call Zebra_A0(Ls_SprB(0), Ls_SprB(1), Ls_SprB(2), Ls_SprB(3), Ls_SprC(6))
                        
                        Case 7  '// 판매가격
                            If Val(Ls_SprC(7)) <> 0 Then '// 증정품인지 Check
                                Call Zebra_A0(Ls_SprB(0), Ls_SprB(1), Ls_SprB(2), Ls_SprB(3), Format(Ls_SprC(7), "###,###,###"))
                                Pri_Sw = 0
                            Else
                                Me.Spr_B.Row = 11   '// 증정품 위치 및 크기
                                Me.Spr_B.Col = 2: Gift_Pr1 = Me.Spr_B.Text
                                Me.Spr_B.Col = 3: Gift_Pr2 = Me.Spr_B.Text
                                Me.Spr_B.Col = 4: Gift_Pr3 = Me.Spr_B.Text
                                Call Label_Print(Ls_SprC(7), Val(Gift_Pr1), Val(Gift_Pr2), Val(Gift_Pr3))
                                Pri_Sw = 1
                            End If

                        Case 8  '// 가격표시-1
                            If Pri_Sw = 0 Then
                                Call Zebra_A0(Ls_SprB(0), Ls_SprB(1), Ls_SprB(2), Ls_SprB(3), Format(Mid(Ls_SprC(7), 1, Len(Trim(Ls_SprC(7))) - 2), "@@@.@"))
                            End If
                        Case 9  '// 가격표시-2
                            If Pri_Sw = 0 Then
                                Call Zebra_A0(Ls_SprB(0), Ls_SprB(1), Ls_SprB(2), Ls_SprB(3), Format(Mid(Ls_SprC(7), 1, Len(Trim(Ls_SprC(7))) - 2), "@@@.@"))
                            End If
                        Case 10 '// 바코드하
                            Call Zebra_BarY(Ls_SprB(0), Ls_SprB(1), Ls_SprB(2), Ls_SprB(3), Ls_SprC(0) & Ls_SprC(1) & Ls_SprC(2) & Ls_SprC(3))
                        Case 11 '// 기타-1
                        Case 12 '// KSP-1
                            Ls_KSP = ""
                            Ls_KSP = Ls_SprB(4) & Ls_SprC(9)
'                            Call Zebra_Han(Val(Ls_SprB(0)), Val(Ls_SprB(1)), Val(Ls_SprB(2)), Val(Ls_SprB(3)), Ls_SprC(9))
                            Call Label_Print(Ls_KSP, Val(Ls_SprB(0)), Val(Ls_SprB(1)), Val(Ls_SprB(2)))
                        Case 13 '// KSP-2
                            Ls_KSP = ""
                            Ls_KSP = Ls_SprB(4) & Ls_SprC(10)
'                            Call Zebra_Han(Val(Ls_SprB(0)), Val(Ls_SprB(1)), Val(Ls_SprB(2)), Val(Ls_SprB(3)), Ls_SprC(10))
                            Call Label_Print(Ls_KSP, Val(Ls_SprB(0)), Val(Ls_SprB(1)), Val(Ls_SprB(2)))
                        Case 14 '// KSP-3
                            Ls_KSP = ""
                            Ls_KSP = Ls_SprB(4) & Ls_SprC(11)
'                            Call Zebra_Han(Val(Ls_SprB(0)), Val(Ls_SprB(1)), Val(Ls_SprB(2)), Val(Ls_SprB(3)), Ls_SprC(11))
                            Call Label_Print(Ls_KSP, Val(Ls_SprB(0)), Val(Ls_SprB(1)), Val(Ls_SprB(2)))
                        Case 15 '// KSP-4
                            Ls_KSP = ""
                            Ls_KSP = Ls_SprB(4) & Ls_SprC(12)
'                            Call Zebra_Han(Val(Ls_SprB(0)), Val(Ls_SprB(1)), Val(Ls_SprB(2)), Val(Ls_SprB(3)), Ls_SprC(12))
                            Call Label_Print(Ls_KSP, Val(Ls_SprB(0)), Val(Ls_SprB(1)), Val(Ls_SprB(2)))
                    End Select
                End If
            Next j
            
'            Frm_Main.Mcom.Output = "^CI0"
            Frm_Main.Mcom.Output = "^PQ" & Ls_SprC(8) & ",0,1,Y^XZ"
'            Frm_Main.Mcom.Output = "~HS"
'            Frm_Main.Mcom.Output = "^XA^IDR:*.*^XZ"
            
        End If
        
    Next i

'  Frm_Main.Mcom.PortOpen = False
  MousePointer = 0
  Cmd_Printer.Enabled = True
End Sub

'***********************************************************************************
'***  Description   :  색상 코드값으로 한글 찾기
'***  Modification Log : 2007/06/14  김동호  Initial Coding
'***********************************************************************************
Private Function Color_Shr(Col_Tmp As String) As String
    Dim i As Integer
    
        For i = 1 To Frm_Setting.Spr_Setting.MaxRows Step 1
            Frm_Setting.Spr_Setting.Row = i
            Frm_Setting.Spr_Setting.Col = 1
            
            If Trim(Col_Tmp) = Frm_Setting.Spr_Setting.Text Then
                Frm_Setting.Spr_Setting.Col = 2
                Color_Shr = Frm_Setting.Spr_Setting.Text
            End If
        Next i
        
        Frm_Setting.Visible = False

End Function

'***********************************************************************************
'***  Description   :  Zebra Start
'***  Modification Log : 2007/06/14  김동호  Initial Coding
'***********************************************************************************
Private Sub Zebra_sTr()
    Frm_Main.Mcom.Output = "^XA^MUd^POI^FS"
'    Frm_Main.Mcom.Output = "^MD" & Cbo_HeadDarkness.Text & "^FS"
'    Frm_Main.Mcom.Output = "^PR" & Trim(Cbo_PrinterSpeed.Text) & "^FS"
    Frm_Main.Mcom.Output = "^LH" & Txt_CenterX.Text & "," & Txt_CenterY.Text & "^FS"
'    Frm_Main.Mcom.Output = "^SEE:UHANGUL.DAT^FS"
'    Frm_Main.Mcom.Output = "^CWQ,E:IDIF.FNT^FS"
End Sub

'***********************************************************************************
'***  Description   :  Zebra A0를 이용한 영문 Set
'***  Modification Log : 2007/06/14  김동호  Initial Coding
'***********************************************************************************
Private Sub Zebra_A0(FoX As String, FoY As String, A0X As String, A0Y As String, FDData As String)
'    Frm_Main.Mcom.Output = "^CI0"
    Frm_Main.Mcom.Output = "^FO" & FoX & "," & FoY
    Frm_Main.Mcom.Output = "^A0," & A0X & "," & A0Y
    Frm_Main.Mcom.Output = "^FD" & FDData & "^FS"
End Sub

'***********************************************************************************
'***  Description   :  Zebra ^CI26을 이용한 한글인쇄
'***  Modification Log : 2007/06/14  김동호  Initial Coding
'***********************************************************************************
Private Sub Zebra_Han(FoHX As String, FoHY As String, A0HX As String, A0HY As String, FDHData As String)
    Frm_Main.Mcom.Output = "^CI26"
    Frm_Main.Mcom.Output = "^FO" & FoHX & "," & FoHY
    Frm_Main.Mcom.Output = "^AQ," & A0HX & "," & A0HY
    Frm_Main.Mcom.Output = "^FD" & FDHData & "^FS"
End Sub

'***********************************************************************************
'***  Description   :  Zebra ^BC를 이용한 바코드 인쇄
'***  Modification Log : 2007/06/14  김동호  Initial Coding
'***********************************************************************************
Private Sub Zebra_BarY(FbX As String, FbY As String, BcX As String, BcY As String, FDBdata As String)
'    Frm_Main.Mcom.Output = "^CI0"
    Frm_Main.Mcom.Output = "^FO" & FbX & "," & FbY
    Frm_Main.Mcom.Output = "^BY" & BcX
    Frm_Main.Mcom.Output = "^BCN," & BcY & ",Y,N,N"
    Frm_Main.Mcom.Output = "^FD" & FDBdata & "^FS"
End Sub

'***********************************************************************************
'***  Description   :  Zebra ^BC를 이용한 바코드 인쇄
'***  Modification Log : 2007/06/14  김동호  Initial Coding
'***********************************************************************************
Private Sub Zebra_BarN(FbX As String, FbY As String, BcX As String, BcY As String, FDBdata As String)
'    Frm_Main.Mcom.Output = "^CI0"
    Frm_Main.Mcom.Output = "^FO" & FbX & "," & FbY
    Frm_Main.Mcom.Output = "^BY" & BcX
    Frm_Main.Mcom.Output = "^BCN," & BcY & ",N,N,N"
    Frm_Main.Mcom.Output = "^FD" & FDBdata & "^FS"
End Sub

'***********************************************************************************
'***  Description   :  폼 Activate 이벤트 정보
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************
Private Sub Form_Activate()
    Frm_Main.Mun_Save.Enabled = True
    Frm_Main.Mun_Close.Enabled = True
    Frm_Main.Mun_AllClose.Enabled = True
    Frm_Main.Mun_Setting.Enabled = True
    Frm_Main.Mun_View.Enabled = True
    Frm_Main.Mun_Windows.Enabled = True
    Frm_Main.tlbMain.Buttons(4).Enabled = True
    
    If Frm_Main.Mcom.PortOpen = True Then
        Frm_Main.Mcom.PortOpen = False
    End If
    
    Frm_Main.Mcom.CommPort = Trim(Right(Me.cbo_Port, 1))
    Frm_Main.Mcom.Settings = Trim(Me.Cbo_Baud) & ",n,8,1"
    Frm_Main.Mcom.RThreshold = 1
    Frm_Main.Mcom.PortOpen = True

End Sub

'***********************************************************************************
'***  Description   :  KSP 선택된 Col 넓이 넓히기
'***  Modification Log : 2007/06/14  김동호  Initial Coding
'***********************************************************************************
Private Sub Spr_CSet(Colno As Long)
    
    If Colno <= 10 Then Exit Sub

    With Spr_C
'        .Col = 1:   .ColWidth(1) = 2.5
'        .Col = 2:   .ColWidth(2) = 4
'        .Col = 3:   .ColWidth(3) = 4
'        .Col = 4:   .ColWidth(4) = 4
'        .Col = 5:   .ColWidth(5) = 4
'        .Col = 6:   .ColWidth(6) = 4
'        .Col = 7:   .ColWidth(7) = 14
'        .Col = 8:   .ColWidth(8) = 5
'        .Col = 9:   .ColWidth(9) = 5
'        .Col = 10:   .ColWidth(10) = 3
        .Col = 11:   .ColWidth(11) = 3
        .Col = 12:   .ColWidth(12) = 3
        .Col = 13:   .ColWidth(13) = 3
        .Col = 14:   .ColWidth(14) = 3
        .Col = Colno: .ColWidth(Colno) = .ColWidth(Colno) * 8.5
        
    End With
End Sub

'***********************************************************************************
'***  Description   :  폼 로드 정보
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************
Private Sub Form_Load()
 
 Dim LS_Filename As String
 

' Frm_Main.Mcom.Output = "~JA"

 
 GS_FromCount = GS_FromCount + 1
 Me.Tag = Str(GS_FromCount)
 
 If CurrentFilename <> "" Then
       
       Me.Caption = CurrentFilename
       CurrentFilename = ""
 Else

       Me.Caption = "새로운 파일"
       
 End If
 
' If Frm_Main.Mcom.PortOpen = True Then
'    Frm_Main.Mcom.PortOpen = False
'End If
'
'Frm_Main.Mcom.CommPort = 2
'Frm_Main.Mcom.Settings = Trim(Me.Cbo_Baud) & ",n,8,1"
'Frm_Main.Mcom.RThreshold = 1
'Frm_Main.Mcom.PortOpen = True

End Sub

'***********************************************************************************
'***  Description   :  폼 언로드 정보
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************

Private Sub Form_Unload(Cancel As Integer)
 
 GS_FromCount = GS_FromCount - 1

 If GS_FromCount = 0 Then
       
       Frm_Main.Mun_Save.Enabled = False
       Frm_Main.Mun_Close.Enabled = False
       Frm_Main.Mun_AllClose.Enabled = False
       Frm_Main.Mun_Setting.Enabled = False
       Frm_Main.Mun_View.Enabled = False
       Frm_Main.Mun_Windows.Enabled = False
       Frm_Main.tlbMain.Buttons(4).Enabled = False
       CurrentFilename = ""
 
 End If

End Sub

'/// -------------------------------------------------------------------///
'/// ---------------++   스프레드 원 클릭 체크   ++ --------------------///
'/// -------------------------------------------------------------------///
Private Sub Spr_C_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Call GfxSelectClear(Spr_C)
    Call G_SET_SpreadColorChange(Spr_C, 1, NewRow)
    Call Spr_CSet(NewCol)
End Sub

'/// -------------------------------------------------------------------///
'/// ----------------++  스프레드 원클릭 체크함  ++ --------------------///
'/// -------------------------------------------------------------------///
Private Sub Spr_B_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Call GfxSelectClear(Spr_B)
    Call G_SET_SpreadColorChange(Spr_B, 1, NewRow)
    Call Spr_CSet(NewCol)
End Sub

'***********************************************************************************
'***  Description   : TabStrip 이벤트 정보
'***  Modification Log : 2006/03/20  김동후  Initial Coding
'***********************************************************************************

Private Sub TabStrip_Click()
  
 If TabStrip.SelectedItem.Index = 1 Then
      
      Fam_B.Visible = True
      Fam_C.Visible = False
 
 ElseIf TabStrip.SelectedItem.Index = 2 Then
      
      Fam_B.Visible = False
      Fam_C.Visible = True
 
 End If

End Sub

