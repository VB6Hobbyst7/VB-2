VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{48932A52-981F-101B-A7FB-4A79242FD97B}#2.0#0"; "TAB32X20.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmReport 
   Caption         =   "결과출력"
   ClientHeight    =   7530
   ClientLeft      =   405
   ClientTop       =   2625
   ClientWidth     =   12435
   ControlBox      =   0   'False
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
   MDIChild        =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   12435
   Begin VB.TextBox txtToJeobsuT2 
      Height          =   330
      Left            =   3600
      MaxLength       =   2
      TabIndex        =   37
      Top             =   1125
      Width           =   330
   End
   Begin VB.TextBox txtToJeobsuT1 
      Height          =   330
      Left            =   3285
      MaxLength       =   2
      TabIndex        =   36
      Top             =   1125
      Width           =   330
   End
   Begin VB.TextBox txtFrJeobsuT2 
      Height          =   330
      Left            =   3600
      MaxLength       =   2
      TabIndex        =   35
      Top             =   765
      Width           =   330
   End
   Begin VB.TextBox txtFrJeobsuT1 
      Height          =   330
      Left            =   3285
      MaxLength       =   2
      TabIndex        =   34
      Top             =   765
      Width           =   330
   End
   Begin Threed.SSCommand cmdSelect 
      Height          =   375
      Left            =   540
      TabIndex        =   28
      Top             =   1890
      Width           =   870
      _Version        =   65536
      _ExtentX        =   1535
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "전체선택"
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdPr 
      Height          =   1050
      Left            =   540
      TabIndex        =   26
      Top             =   2340
      Width           =   870
      _Version        =   65536
      _ExtentX        =   1535
      _ExtentY        =   1852
      _StockProps     =   78
      Caption         =   "결과출력"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Outline         =   0   'False
   End
   Begin TabproLib.vaTabPro vaTabPro1 
      Height          =   1140
      Left            =   5040
      OleObjectBlob   =   "frmReport.frx":0000
      TabIndex        =   10
      Top             =   420
      Width           =   6810
      Begin VB.TextBox txtToLabno 
         Enabled         =   0   'False
         Height          =   330
         Left            =   -18749
         TabIndex        =   33
         Text            =   "txtToLabno"
         Top             =   -15824
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "검사일자"
         Enabled         =   0   'False
         Height          =   180
         Left            =   -19499
         TabIndex        =   32
         Top             =   -16109
         Width           =   1050
      End
      Begin VB.OptionButton Option1 
         Caption         =   "접수일자"
         Enabled         =   0   'False
         Height          =   180
         Left            =   -18374
         TabIndex        =   31
         Top             =   -16109
         Value           =   -1  'True
         Width           =   1050
      End
      Begin VB.ComboBox cmbSample 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -18704
         Style           =   2  '드롭다운 목록
         TabIndex        =   25
         Top             =   -15839
         Width           =   2625
      End
      Begin VB.ComboBox cmbUser 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -18344
         Style           =   2  '드롭다운 목록
         TabIndex        =   24
         Top             =   -15839
         Width           =   1545
      End
      Begin VB.TextBox txtPtno 
         Enabled         =   0   'False
         Height          =   330
         Left            =   -17489
         TabIndex        =   15
         Text            =   "txtPtno"
         Top             =   -15869
         Width           =   1680
      End
      Begin VB.TextBox txtSname 
         Enabled         =   0   'False
         Height          =   330
         Left            =   -18389
         TabIndex        =   14
         Text            =   "txtSname"
         Top             =   -15869
         Width           =   1725
      End
      Begin VB.TextBox txtLabno 
         Enabled         =   0   'False
         Height          =   330
         Left            =   -17579
         TabIndex        =   13
         Text            =   "txtLabno"
         Top             =   -15824
         Width           =   1005
      End
      Begin VB.ComboBox cmbDept 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -18389
         Style           =   2  '드롭다운 목록
         TabIndex        =   12
         Top             =   -15839
         Width           =   1860
      End
      Begin VB.ComboBox cmbWard 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -19019
         Style           =   2  '드롭다운 목록
         TabIndex        =   11
         Top             =   -15794
         Width           =   2355
      End
      Begin VB.Label Label3 
         Caption         =   "왼쪽에 있는 Date 의 조건으로 조회합니다"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -17159
         TabIndex        =   30
         Top             =   -16214
         Width           =   1950
      End
      Begin MSForms.CommandButton cmdQuery7 
         Height          =   510
         Left            =   -21044
         TabIndex        =   23
         Top             =   -15959
         Width           =   1680
         VariousPropertyBits=   25
         Caption         =   "조회확인"
         PicturePosition =   327683
         Size            =   "2963;900"
         FontName        =   "굴림체"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdQuery6 
         Height          =   510
         Left            =   -21089
         TabIndex        =   22
         Top             =   -15959
         Width           =   1680
         VariousPropertyBits=   25
         Caption         =   "조회확인"
         PicturePosition =   327683
         Size            =   "2963;900"
         FontName        =   "굴림체"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdQuery5 
         Height          =   510
         Left            =   -21089
         TabIndex        =   21
         Top             =   -15959
         Width           =   1680
         VariousPropertyBits=   25
         Caption         =   "조회확인"
         PicturePosition =   327683
         Size            =   "2963;900"
         FontName        =   "굴림체"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdQuery4 
         Height          =   510
         Left            =   -21089
         TabIndex        =   20
         Top             =   -15959
         Width           =   1680
         VariousPropertyBits=   25
         Caption         =   "조회확인"
         PicturePosition =   327683
         Size            =   "2963;900"
         FontName        =   "굴림체"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdQuery3 
         Height          =   510
         Left            =   -21089
         TabIndex        =   19
         Top             =   -15959
         Width           =   1680
         VariousPropertyBits=   25
         Caption         =   "조회확인"
         PicturePosition =   327683
         Size            =   "2963;900"
         FontName        =   "굴림체"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdQuery2 
         Height          =   510
         Left            =   -21089
         TabIndex        =   18
         Top             =   -15959
         Width           =   1680
         VariousPropertyBits=   25
         Caption         =   "조회확인"
         PicturePosition =   327683
         Size            =   "2963;900"
         FontName        =   "굴림체"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdQuery1 
         Height          =   510
         Left            =   -21089
         TabIndex        =   17
         Top             =   -15959
         Width           =   1680
         VariousPropertyBits=   25
         Caption         =   "조회확인"
         PicturePosition =   327683
         Size            =   "2963;900"
         FontName        =   "굴림체"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdQuery 
         Height          =   510
         Left            =   -21074
         TabIndex        =   16
         Top             =   -16304
         Width           =   1455
         VariousPropertyBits=   25
         Caption         =   "조회확인"
         PicturePosition =   327683
         Size            =   "2566;900"
         FontName        =   "굴림체"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
   End
   Begin VB.OptionButton optWhere 
      Caption         =   "결과완료"
      Height          =   285
      Index           =   2
      Left            =   3870
      TabIndex        =   8
      Top             =   1530
      Value           =   -1  'True
      Width           =   1050
   End
   Begin VB.OptionButton optWhere 
      Caption         =   "부분결과"
      Height          =   285
      Index           =   1
      Left            =   2790
      TabIndex        =   7
      Top             =   1530
      Width           =   1050
   End
   Begin VB.OptionButton optWhere 
      Caption         =   "접수중"
      Height          =   285
      Index           =   0
      Left            =   1890
      TabIndex        =   6
      Top             =   1530
      Width           =   870
   End
   Begin VB.ComboBox cmbSLip 
      Height          =   300
      Left            =   1890
      Style           =   2  '드롭다운 목록
      TabIndex        =   4
      Top             =   405
      Width           =   2805
   End
   Begin MSComCtl2.DTPicker dtFrDate 
      Height          =   330
      Left            =   1890
      TabIndex        =   2
      Top             =   765
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   24772611
      CurrentDate     =   36508
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   180
      Top             =   4635
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":0436
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":0752
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '위 맞춤
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12435
      _ExtentX        =   21934
      _ExtentY        =   635
      ButtonWidth     =   1508
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "Exit"
            Description     =   "Exit of PrintForm"
            Object.ToolTipText     =   "Exit of PrintForm"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clear"
            Key             =   "Clear"
            Description     =   "Clear of Form"
            Object.ToolTipText     =   "Clear of Form"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin FPSpreadADO.fpSpread ssResult 
      Height          =   5460
      Left            =   1560
      TabIndex        =   1
      Top             =   1860
      Width           =   10335
      _Version        =   196608
      _ExtentX        =   18230
      _ExtentY        =   9631
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   7
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   18
      MaxRows         =   600
      Protect         =   0   'False
      ScrollBars      =   2
      ShadowColor     =   12632256
      ShadowDark      =   8421504
      ShadowText      =   0
      SpreadDesigner  =   "frmReport.frx":0A6E
      UserResize      =   0
      VisibleCols     =   18
      VisibleRows     =   500
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker dtToDate 
      Height          =   330
      Left            =   1890
      TabIndex        =   3
      Top             =   1125
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   24772611
      CurrentDate     =   36508
   End
   Begin MSForms.CommandButton cmdPrMicro 
      Height          =   555
      Left            =   585
      TabIndex        =   38
      Top             =   3555
      Visible         =   0   'False
      Width           =   690
      Caption         =   "42"
      Size            =   "1217;979"
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000007&
      Caption         =   "Label5"
      Height          =   375
      Left            =   585
      TabIndex        =   29
      Top             =   1935
      Width           =   870
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      Caption         =   "Label4"
      Height          =   1050
      Left            =   630
      TabIndex        =   27
      Top             =   2385
      Width           =   825
   End
   Begin VB.Label Label2 
      Caption         =   "출력검사종목?:"
      Height          =   240
      Left            =   180
      TabIndex        =   9
      Top             =   450
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "접수일자:From/To"
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   810
      Width           =   1590
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Function Printing_RCode(ByVal sJeobsuDt As String, ByVal iSLno1 As Integer, ByVal iSLno2 As Integer, ByVal sItemCd As String) As String
    Dim adoResult       As ADODB.Recordset
    Dim sRcode(1 To 5)  As String
    Dim sResult(1 To 5) As String
    
    
    strSql = ""
    strSql = strSql & " SELECT a.*"
    strSql = strSql & " FROM   TWEXAM_GENERAL_SUB a"
    strSql = strSql & " WHERE   a.JEOBSUDT  = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND     a.SLipno1   = " & iSLno1
    strSql = strSql & " AND     a.SLipno2   = " & iSLno2
    strSql = strSql & " AND     a.ItemCd    = '" & sItemCd & "'"
    
    
    If adoSetOpen(strSql, adoResult) Then
        sRcode(1) = adoResult("Rcode1").Value & "": sResult(1) = adoResult("Result1").Value & ""
        sRcode(2) = adoResult("Rcode2").Value & "": sResult(2) = adoResult("Result2").Value & ""
        sRcode(3) = adoResult("Rcode3").Value & "": sResult(3) = adoResult("Result3").Value & ""
        sRcode(4) = adoResult("Rcode4").Value & "": sResult(4) = adoResult("Result4").Value & ""
        sRcode(5) = adoResult("Rcode5").Value & "": sResult(5) = adoResult("Result5").Value & ""
                
        If Trim(sRcode(1)) <> "" Or Trim(sResult(1)) <> "" Then
            Printing_RCode = "@." & _
                              Get_OrgName(sRcode(1)) & "  -" & Trim(sResult(1)): End If
        
        If Trim(sRcode(2)) <> "" Or Trim(sResult(2)) <> "" Then
            Printing_RCode = "@." & _
                             Printing_RCode & vbCrLf & _
                             Get_OrgName(sRcode(2)) & "  -" & Trim(sResult(2)): End If
                             
        If Trim(sRcode(3)) <> "" Or Trim(sResult(3)) <> "" Then
            Printing_RCode = "@." & _
                             Printing_RCode & vbCrLf & _
                             Get_OrgName(sRcode(3)) & "  -" & Trim(sResult(3)): End If
                             
        If Trim(sRcode(4)) <> "" Or Trim(sResult(4)) <> "" Then
            Printing_RCode = "@." & _
                             Printing_RCode & vbCrLf & _
                             Get_OrgName(sRcode(4)) & "  -" & Trim(sResult(4)): End If
                             
        If Trim(sRcode(5)) <> "" Or Trim(sResult(5)) <> "" Then
            Printing_RCode = "@." & _
                             Printing_RCode & vbCrLf & _
                             Get_OrgName(sRcode(5)) & "  -" & Trim(sResult(5)): End If
                             
        Call adoSetClose(adoResult)
        
    End If

    
End Function
Public Function Printing_Sens(ByVal sJeobsuDt As String, ByVal iSLno1 As Integer, ByVal iSLno2 As Integer, ByVal sItemCd As String) As Integer
    Dim adoSensRet      As ADODB.Recordset
    Dim iSensCls        As Integer
    
    '420401 = Drug Senstivity (세균검사)
    
    Call SensResultClear
    
    strSql = ""
    strSql = strSql & "  SELECT  a.ItemCd, c.Org_name, d.Codenm AntiName, a.Sens, a.Value, b.Result1 Result"
    strSql = strSql & "  FROM    TWEXAM_SENS        a,"
    strSql = strSql & "          TWEXAM_GENERAL_Sub b,"
    strSql = strSql & "          TWEXAM_ORGLIST     c,"
    strSql = strSql & "          TWEXAM_AntiList    d"
    strSql = strSql & "  WHERE   a.JEOBSUDT  = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & "  AND     a.SLipno1   = " & iSLno1
    strSql = strSql & "  AND     a.SLipno2   = " & iSLno2
    strSql = strSql & "  AND     a.ItemCd    = '" & sItemCd & "'"
    strSql = strSql & "  AND     a.JeobsuDt = b.JeobsuDt(+)"
    strSql = strSql & "  AND     a.SLipno1  = b.SLipno1(+)"
    strSql = strSql & "  AND     a.SLipno2  = b.SLipno2(+) "
    strSql = strSql & "  AND     a.ORACOD   = b.Rcode1"
    strSql = strSql & "  AND     a.Oracod   = c.Org_code(+)"
    strSql = strSql & "  AND     a.Yakcod   = d.Codeky(+)"
    
    strSql = strSql & " UNION ALL "
    strSql = strSql & "  SELECT  a.ItemCd, c.Org_name, d.Codenm AntiName, a.Sens, a.Value, b.Result2 Result"
    strSql = strSql & "  FROM    TWEXAM_SENS        a,"
    strSql = strSql & "          TWEXAM_GENERAL_Sub b,"
    strSql = strSql & "          TWEXAM_ORGLIST     c,"
    strSql = strSql & "          TWEXAM_AntiList    d"
    strSql = strSql & "  WHERE   a.JEOBSUDT  = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & "  AND     a.SLipno1   = " & iSLno1
    strSql = strSql & "  AND     a.SLipno2   = " & iSLno2
    strSql = strSql & "  AND     a.ItemCd    = '" & sItemCd & "'"
    strSql = strSql & "  AND     a.JeobsuDt = b.JeobsuDt(+)"
    strSql = strSql & "  AND     a.SLipno1  = b.SLipno1(+)"
    strSql = strSql & "  AND     a.SLipno2  = b.SLipno2(+) "
    strSql = strSql & "  AND     a.ORACOD   = b.Rcode2"
    strSql = strSql & "  AND     a.Oracod   = c.Org_code(+)"
    strSql = strSql & "  AND     a.Yakcod   = d.Codeky(+)"
    
    strSql = strSql & " UNION ALL "
    strSql = strSql & "  SELECT  a.ItemCd, c.Org_name, d.Codenm AntiName, a.Sens, a.Value, b.Result3 Result"
    strSql = strSql & "  FROM    TWEXAM_SENS        a,"
    strSql = strSql & "          TWEXAM_GENERAL_Sub b,"
    strSql = strSql & "          TWEXAM_ORGLIST     c,"
    strSql = strSql & "          TWEXAM_AntiList    d"
    strSql = strSql & "  WHERE   a.JEOBSUDT  = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & "  AND     a.SLipno1   = " & iSLno1
    strSql = strSql & "  AND     a.SLipno2   = " & iSLno2
    strSql = strSql & "  AND     a.ItemCd    = '" & sItemCd & "'"
    strSql = strSql & "  AND     a.JeobsuDt = b.JeobsuDt(+)"
    strSql = strSql & "  AND     a.SLipno1  = b.SLipno1(+)"
    strSql = strSql & "  AND     a.SLipno2  = b.SLipno2(+) "
    strSql = strSql & "  AND     a.ORACOD   = b.Rcode3"
    strSql = strSql & "  AND     a.Oracod   = c.Org_code(+)"
    strSql = strSql & "  AND     a.Yakcod   = d.Codeky(+)"
    
    strSql = strSql & " UNION ALL "
    strSql = strSql & "  SELECT  a.ItemCd, c.Org_name, d.Codenm AntiName, a.Sens, a.Value, b.Result4 Result"
    strSql = strSql & "  FROM    TWEXAM_SENS        a,"
    strSql = strSql & "          TWEXAM_GENERAL_Sub b,"
    strSql = strSql & "          TWEXAM_ORGLIST     c,"
    strSql = strSql & "          TWEXAM_AntiList    d"
    strSql = strSql & "  WHERE   a.JEOBSUDT  = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & "  AND     a.SLipno1   = " & iSLno1
    strSql = strSql & "  AND     a.SLipno2   = " & iSLno2
    strSql = strSql & "  AND     a.ItemCd    = '" & sItemCd & "'"
    strSql = strSql & "  AND     a.JeobsuDt = b.JeobsuDt(+)"
    strSql = strSql & "  AND     a.SLipno1  = b.SLipno1(+)"
    strSql = strSql & "  AND     a.SLipno2  = b.SLipno2(+) "
    strSql = strSql & "  AND     a.ORACOD   = b.Rcode4"
    strSql = strSql & "  AND     a.Oracod   = c.Org_code(+)"
    strSql = strSql & "  AND     a.Yakcod   = d.Codeky(+)"
    
    strSql = strSql & " UNION ALL "
    strSql = strSql & "  SELECT  a.ItemCd, c.Org_name, d.Codenm AntiName, a.Sens, a.Value, b.Result5 Result"
    strSql = strSql & "  FROM    TWEXAM_SENS        a,"
    strSql = strSql & "          TWEXAM_GENERAL_Sub b,"
    strSql = strSql & "          TWEXAM_ORGLIST     c,"
    strSql = strSql & "          TWEXAM_AntiList    d"
    strSql = strSql & "  WHERE   a.JEOBSUDT  = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & "  AND     a.SLipno1   = " & iSLno1
    strSql = strSql & "  AND     a.SLipno2   = " & iSLno2
    strSql = strSql & "  AND     a.ItemCd    = '" & sItemCd & "'"
    strSql = strSql & "  AND     a.JeobsuDt = b.JeobsuDt(+)"
    strSql = strSql & "  AND     a.SLipno1  = b.SLipno1(+)"
    strSql = strSql & "  AND     a.SLipno2  = b.SLipno2(+) "
    strSql = strSql & "  AND     a.ORACOD   = b.Rcode5"
    strSql = strSql & "  AND     a.Oracod   = c.Org_code(+)"
    strSql = strSql & "  AND     a.Yakcod   = d.Codeky(+)"
    
    If False = adoSetOpen(strSql, adoSensRet) Then Exit Function
    
    For iSensCls = 0 To adoSensRet.RecordCount - 1
        SensResult.ItemCd(iSensCls) = adoSensRet.Fields("ItemCd").Value & ""
        SensResult.Rcode(iSensCls) = adoSensRet.Fields("Org_name").Value & ""
        SensResult.AntiName(iSensCls) = adoSensRet.Fields("AntiName").Value & ""
        SensResult.Sens(iSensCls) = adoSensRet.Fields("Sens").Value & ""
        SensResult.Value(iSensCls) = adoSensRet.Fields("Value").Value & ""
        SensResult.Value(iSensCls) = adoSensRet.Fields("Result").Value & ""
        adoSensRet.MoveNext
    Next
    Printing_Sens = adoSensRet.RecordCount
    Call adoSetClose(adoSensRet)
    
    
    
End Function

Public Function GET_General_Status(ByVal sJeobsuDt As String, ByVal iSLno1 As Integer, ByVal iSLno2 As Integer) As String
    Dim adoStatus       As ADODB.Recordset
    
    strSql = ""
    strSql = strSql & " SELECT Status"
    strSql = strSql & " FROM   TWEXAM_General"
    strSql = strSql & " WHERE  JEOBSUDT  = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    SLipno1   = " & iSLno1
    strSql = strSql & " AND    SLipno2   = " & iSLno2
    
    If False = adoSetOpen(strSql, adoStatus) Then
        GET_General_Status = ""
        Exit Function
    End If
    
    Select Case adoStatus.Fields("Status").Value & ""
        Case "R": GET_General_Status = "접수중"
        Case "C":
            If iSLno1 = 42 Then
                GET_General_Status = "최종보고"
            Else
                GET_General_Status = "결과완료"
            End If
        Case "P":
            If iSLno1 = 42 Then
                GET_General_Status = "예비보고"
            Else
                GET_General_Status = "부분결과확인"
            End If
        Case "U": GET_General_Status = "미확인"
        Case "X": GET_General_Status = "Data이상(Panic Or Delta)"
        Case Else: GET_General_Status = ""
    End Select
    
    Call adoSetClose(adoStatus)

End Function

    
    

Private Sub cmdPr_Click()
    Dim sSLipno1    As String
    Dim sSLipno2    As String
    Dim sPtno       As String * 8
    Dim sJeobsuDt   As String
    Dim sSex        As String
    Dim sGeomsaDt   As String
    Dim sAge        As String
    Dim sRemark     As String
    Dim sSname      As String * 10
    Dim sRitemCd    As String * 8
    Dim sRitemNm    As String * 30
    Dim sDanWi      As String * 6
    Dim sResult     As String * 12
    Dim sResult42   As String
    Dim sMin        As String * 6
    Dim sMax        As String * 6
    Dim sRoomCode   As String
    Dim sDeptName   As String
    Dim sSamplename As String
    Dim iSensCount      As Integer
    Dim iSensOrderName  As String
    
    If Left(cmbSLip.Text, 2) = "42" Then
        Call cmdPrMicro_Click
        Exit Sub
    End If
    
    If vbNo = MsgBox(" Print 하시겠습니까?", vbQuestion + vbYesNo, "Printing 할까요?") Then
        Exit Sub
    End If
    
    '서울 건양병원(전산실)에서 일괄 Printing 하기위하여 Sort한 Sub 임.
    '필요할시 Comment 제거후 쓸것(과별,환자번호,검사종목(slipno1)
    'GoSub PrintSpread_Sort

    Printer.Orientation = vbPRORPortrait
    'Printer.Orientation = vbPRORLandscape

    For i = 1 To ssResult.DataRowCnt
        ssResult.Row = i
        ssResult.Col = 1
            
        'If Left(cmbSLip.Text, 2) = "15" Then
        '    MsgBox "골수검사는 결과입력Form 에서 개별로 출력이 가능합니다!..."
        '    Exit Sub
        'Else
            If ssResult.Value = True Then
                ssResult.Col = 2: sJeobsuDt = ssResult.Text
                ssResult.Col = 3: sSLipno1 = ssResult.Text
                ssResult.Col = 4: sSLipno2 = ssResult.Text
                ssResult.Col = 6: sPtno = ssResult.Text
                GoSub Main_Process
            End If
        'End If
    Next
    Exit Sub
    
PrintSpread_Sort:
    ssResult.Row = 1
    ssResult.Row2 = ssResult.DataRowCnt
    ssResult.Col = 1
    ssResult.Col2 = ssResult.DataColCnt
    
    ssResult.SortBy = SS_SORT_BY_ROW
    ssResult.SortKey(1) = 14        '과별
    ssResult.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
    
    ssResult.SortKey(2) = 6        '환자번호
    ssResult.SortKeyOrder(2) = SS_SORT_ORDER_ASCENDING
    
    ssResult.SortKey(3) = 5         '검사종목
    ssResult.SortKeyOrder(5) = SS_SORT_ORDER_ASCENDING

    ssResult.Action = SS_ACTION_SORT

    
    Return
    
    
Main_Process:
    
    strSql = ""
    strSql = strSql & " SELECT TO_CHAR(a.JeobsuDt, 'yyyy-MM-dd') JeobsuDt,"
    strSql = strSql & "        a.SLipno1, a.SLipno2, a.Ptno, b.Sname, b.Sex, a.AgeYY,"
    strSql = strSql & "        TO_CHAR(a.GeomsaDt,'yyyy-MM-dd') GeomsaDt,"
    strSql = strSql & "        a.RoomCode, a.GeomsaCM, c.DeptNamek, d.WardCode"
    strSql = strSql & " FROM   TWEXAM_GENERAL  a,"
    strSql = strSql & "        TWEXAM_IDNOMST  b,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT      c,"
    strSql = strSql & "        TW_MIS_PMPA.TWBas_ROOM      d "
    strSql = strSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.Slipno1  =  " & Val(sSLipno1)
    strSql = strSql & " AND    a.SLipno2  =  " & Val(sSLipno2)
    strSql = strSql & " AND    a.PTNO     = '" & sPtno & "'"
    strSql = strSql & " AND    a.Ptno     = b.Ptno(+)"
    strSql = strSql & " AND    a.DeptCode = c.DeptCode(+)"
    strSql = strSql & " AND    a.RoomCode = d.RoomCode(+)"
    
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    sJeobsuDt = adoSet.Fields("JeobsuDt").Value & ""
    sSLipno1 = adoSet.Fields("SLipno1").Value & ""
    sSLipno2 = adoSet.Fields("SLipno2").Value & ""
    sPtno = adoSet.Fields("Ptno").Value & ""
    sSname = adoSet.Fields("Sname").Value & ""
    sSex = adoSet.Fields("Sex").Value & ""
    sGeomsaDt = adoSet.Fields("GeomsaDt").Value & ""
    
    sRoomCode = adoSet.Fields("RoomCode").Value & ""
    If Trim(sRoomCode) <> "" Then
        sRoomCode = adoSet.Fields("WardCode").Value & "/" & sRoomCode
    End If
    
    sDeptName = adoSet.Fields("DeptNamek").Value & ""
    sRemark = adoSet.Fields("GeomsaCM").Value & ""
    sAge = adoSet.Fields("AgeYY").Value & ""
    Call adoSetClose(adoSet)
    
    
    GoSub PrintHead_RTN
    GoSub Print_OK_RTN
    Printer.EndDoc
    
    
    
    
    Return


PrintHead_RTN:
    Dim sDeptCode   As String * 10
    Dim sAgeYY      As String
    Dim sJDT        As String
    Dim sGDT        As String
    Dim sSlipTitle  As String
    Dim sGeomsaJa   As String
    Dim sLabno      As String * 5
    Dim adoSpec     As ADODB.Recordset
    Dim adoGen      As ADODB.Recordset
    Dim sStat       As String
    
    
    strSql = ""
    strSql = strSql & " SELECT codenm"
    strSql = strSql & " FROM   TWEXAM_SPECODE"
    strSql = strSql & " WHERE  CODEGU = '12'"
    strSql = strSql & " AND    CODEKY = '" & sSLipno1 & "'"
    If adoSetOpen(strSql, adoSpec) Then
        sSlipTitle = adoSpec.Fields("Codenm").Value & ""
        Call adoSetClose(adoSpec)
    End If
    
'o  strSql = ""
'o  strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_DEPT INDEX_DEPT0) */"

    strSql = ""
    strSql = strSql & " SELECT a.RoomCode, a.Sex, a.AgeYY, a.SLipno1, a.SLipno2, a.Status,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDT,'YYYY-MM-DD') JeobsuDt, a.JeobsuT1, a.JeobsuT2,"
    strSql = strSql & "        TO_CHAR(a.GeomsaDT,'YYYY-MM-DD') GeomsaDt, a.GeomsaT1, a.GeomsaT2,"
    strSql = strSql & "        a.GeomsaCM, b.DeptNamek, c.Name, d.Codenm"
    strSql = strSql & " FROM   TWEXAM_GENERAL a,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT     b,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PASS     c,"
    strSql = strSql & "        TWEXAM_Sample  d "
    strSql = strSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.Ptno     = '" & sPtno & "'"
    strSql = strSql & " AND    a.Slipno1  =  " & Val(sSLipno1)
    strSql = strSql & " AND    a.Slipno2  =  " & Val(sSLipno2)
    strSql = strSql & " AND    a.DeptCode = b.DeptCode(+)"
    strSql = strSql & " AND    a.Geomsaja = c.idNumber(+)"
    strSql = strSql & " AND    a.GeomchCd = d.Code(+)"
    
    If False = adoSetOpen(strSql, adoGen) Then Return
    
    sDeptCode = adoGen.Fields("DeptNamek").Value & ""
    sRoomCode = adoGen.Fields("RoomCode").Value & ""
    sSex = adoGen.Fields("Sex").Value & ""
    sAgeYY = adoGen.Fields("AgeYY").Value & ""
    sJDT = adoGen.Fields("JeobsuDt").Value & " " & adoGen.Fields("JeobsuT1").Value & ":" & adoGen.Fields("JeobsuT2").Value & ""
    sGDT = adoGen.Fields("GeomsaDt").Value & " " & adoGen.Fields("GeomsaT1").Value & ":" & adoGen.Fields("GeomsaT2").Value & ""
    sRemark = Trim(adoGen.Fields("GeomsaCm").Value & "")
    sGeomsaJa = Trim(adoGen.Fields("Name").Value & "")
    sSamplename = Trim(adoGen.Fields("Codenm").Value & "")
    sLabno = adoGen.Fields("SLipno2").Value & ""
    
    Select Case Trim(adoGen.Fields("Status").Value & "")
        Case "R": sStat = "접수중"
        Case "P": sStat = "부분결과"
        Case "U": sStat = "미확인"
        Case "C": sStat = "결과완료"
        Case "X": sStat = "이상Data"
        Case Else: sStat = ""
    End Select
    
    Call adoSetClose(adoGen)
    
    Printer.FontName = "바탕체"
    Printer.FontSize = "12"
    Printer.FontBold = True
    Printer.FontItalic = True
    
    Printer.Print sSlipTitle & "       Result Report   " & sLabno & "  (" & Trim(sStat) & ")"
    Printer.ForeColor = RGB(192, 192, 192)
    Printer.Print "━━━━━━━━━━━━━━━━━━━━━━━━━"
    Printer.ForeColor = RGB(0, 0, 0)
    
    Printer.FontName = "바탕체"
    Printer.FontSize = 9
    Printer.FontBold = False
    Printer.FontItalic = False
    
    Printer.Print "등록No: " & sPtno; Tab(40); "검  체: " & sSamplename
    Printer.Print "성  명: " & Trim(sSname) & "[" & sSex & "/" & sAgeYY & "]"; Tab(40); "검사자: " & sGeomsaJa
    Printer.Print "진료과: " & sDeptName; Tab(40); "접수일: " & sJDT
    Printer.Print "병  실: " & sRoomCode; Tab(40); "검사일: " & sGDT

    If Left(sSLipno1, 1) = "4" Then
        Printer.ForeColor = RGB(192, 192, 192)
        Printer.Print "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
        Printer.ForeColor = RGB(0, 0, 0)
        Printer.Print "       검사항목             :         검사결과                      "
        Printer.ForeColor = RGB(192, 192, 192)
        Printer.Print "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
        Printer.ForeColor = RGB(0, 0, 0)
    Else
        Printer.ForeColor = RGB(192, 192, 192)
        Printer.Print "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
        Printer.ForeColor = RGB(0, 0, 0)
        Printer.Print "       검사항목             :  검사결과 :     참고치      :   단위  "
        Printer.ForeColor = RGB(192, 192, 192)
        Printer.Print "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
        Printer.ForeColor = RGB(0, 0, 0)
    End If
    Return
    
    
    
Print_OK_RTN:
    Dim j               As Integer
    Dim adoGsub         As ADODB.Recordset
    Dim sSensFlag       As String * 1
    Dim nPrcnt          As Integer
    Dim saItemCd(10)    As String
    Dim iL              As Integer
    Dim sFlagAbnormal   As String
    
    sSensFlag = ""
    
    
    strSql = ""
    strSql = strSql & " SELECT TO_CHAR(a.JeobsuDt,'YYYY-MM-DD') JeobsuDt,"
    strSql = strSql & "        a.ItemCd, a.Result1, a.Result2, a.Result3, a.Result4, a.Result5, "
    strSql = strSql & "        b.ItemNM, b.MinCham, b.MaxCham, b.DanWi, b.MinDanger, b.MaxDanger,"
    strSql = strSql & "        b.ResultW"
    strSql = strSql & " FROM   TWEXAM_GENERAL_SUB a,"
    strSql = strSql & "        TWEXAM_ITEMML      b "
    strSql = strSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.Ptno     = '" & sPtno & "'"
    strSql = strSql & " AND    a.Slipno1  =  " & Val(sSLipno1)
    strSql = strSql & " AND    a.Slipno2  =  " & Val(sSLipno2)
    'strSql = strSql & " AND    a.Verify   = 'Y'"
    strSql = strSql & " AND    a.ItemCd   = b.CodeKy(+)"
    strSql = strSql & " ORDER  BY a.itemCd"
    
    If False = adoSetOpen(strSql, adoGsub) Then Return
    
    
    Printer.FontName = "바탕체"
    Printer.FontSize = 9
    Printer.FontBold = False
    Printer.FontItalic = False
    
    For iL = 0 To 10
        saItemCd(iL) = ""
    Next
    
    iL = 0
    Do Until adoGsub.EOF
        sRitemCd = adoGsub.Fields("ItemCd").Value & ""
        sRitemNm = adoGsub.Fields("ItemNm").Value & ""
        RSet sDanWi = adoGsub.Fields("DanWi").Value & ""
        LSet sResult = adoGsub.Fields("Result1").Value & ""
        
        If sSLipno1 <> "42" Then
            GoSub Get_RefData
            If UCase(Trim(adoGsub.Fields("ResultW").Value & "")) = "N" Then
                sFlagAbnormal = "  "
                If Val(sMin) <> 0 Or Val(sMax) <> 0 Then
                    If Val(sResult) < Val(sMin) Then sFlagAbnormal = "↓"
                    If Val(sResult) > Val(sMax) Then sFlagAbnormal = "↑"
                Else
                    sFlagAbnormal = ""        '참고치 값이 Setting 되지 아니한 것은 Marking 없앨것
                End If
            Else
                sFlagAbnormal = ""
            End If
            
            GoSub BanZul_Sub
            
            If UCase(Trim(adoGsub.Fields("ResultW").Value & "")) = "C" Then
                Printer.Print "  " & sRitemNm & sResult & sFlagAbnormal & sMax & sDanWi
            Else
                Printer.Print "  " & sRitemNm & sResult & sFlagAbnormal & sMin & "~ " & sMax & sDanWi
            End If
        Else
            If Trim(adoGsub.Fields("ResultW").Value & "") = "S" Then
                sSensFlag = "*"
                sResult42 = "다음장참조"
                iSensOrderName = "": iSensOrderName = sRitemNm
                saItemCd(iL) = sRitemCd
                iL = iL + 1
                Printer.Print "  " & sRitemNm & sResult42
            Else
                sResult42 = adoGsub.Fields("Result1").Value & ""
                Printer.Print "  " & sRitemNm & Trim(sResult42)
            End If
        End If
        
        adoGsub.MoveNext
    Loop
    Call adoSetClose(adoGsub)
    
    Printer.Print ""
    Printer.Print ""
    Printer.Print "☞ Remark ___________________________"
    Printer.Print ""
    Printer.Print " :" & sRemark
    Printer.NewPage

    Dim sTmpRcode       As String
    Dim sDispRname      As String * 25
    Dim iCount          As Integer
    Dim iPrPageCnt      As Integer
    
    
    
    For iL = 0 To 10
        iCount = iL
        If saItemCd(iL) = "" Then
            Exit For
        End If
    Next
    
    If sSensFlag = "*" Then
        For iL = 0 To iCount
            Printer.FontName = "바탕체"
            Printer.FontSize = "12"
            Printer.FontBold = True
            Printer.FontItalic = True
        
            Printer.Print sSlipTitle & "(SENSTIVITY)  Result Report   " & sLabno & "(" & Trim(sStat) & ")"
            Printer.FontName = "바탕체"
            Printer.FontSize = 9
            Printer.FontBold = False
            Printer.FontItalic = False
            
            Printer.ForeColor = RGB(192, 192, 192)
            Printer.Print "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
            Printer.ForeColor = RGB(0, 0, 0)
            Printer.Print "등록No: " & sPtno; Tab(40); "검  체: " & sSamplename
            Printer.Print "성  명: " & Trim(sSname) & "[" & sSex & "/" & sAgeYY & "]"; Tab(40); "검사자: " & sGeomsaJa
            Printer.Print "진료과: " & sDeptName; Tab(40); "접수일: " & sJDT
            Printer.Print "병  실: " & sRoomCode; Tab(40); "검사일: " & sGDT
            Printer.ForeColor = RGB(192, 192, 192)
            Printer.Print "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
            Printer.ForeColor = RGB(0, 0, 0)
            
            Printer.FontName = "바탕체": Printer.FontSize = 10:  Printer.FontBold = True: Printer.FontItalic = False
            Printer.Print "Order    : " & Get_iTemName(saItemCd(iL))
            Printer.Print "보고형태 : " & GET_General_Status(sJeobsuDt, Val(sSLipno1), Val(sSLipno2))
            
            
            iSensCount = Printing_Sens(sJeobsuDt, Val(sSLipno1), Val(sSLipno2), saItemCd(iL))
                        
            If iSensCount > 0 Then
                For nPrcnt = 0 To iSensCount - 1
                    If sTmpRcode = SensResult.Rcode(nPrcnt) Then
                        Printer.Print "     " & SensResult.AntiName(nPrcnt) & "," & SensResult.Sens(nPrcnt)
                                      
                    Else
                        Printer.Print ""
                        Printer.FontName = "바탕체": Printer.FontSize = 9:  Printer.FontBold = True: Printer.FontItalic = False
                        Printer.Print "@." & _
                                      SensResult.Rcode(nPrcnt) & " " & SensResult.Result(nPrcnt) & vbCrLf & _
                                      "     " & SensResult.AntiName(nPrcnt) & "," & SensResult.Sens(nPrcnt)
                    End If
                    sTmpRcode = SensResult.Rcode(nPrcnt)
                Next
            Else
                If Trim(saItemCd(iL)) <> "" Then
                    Printer.Print ""
                    Printer.FontName = "바탕체": Printer.FontSize = 9:  Printer.FontBold = True: Printer.FontItalic = False
                    Printer.Print Printing_RCode(sJeobsuDt, Val(sSLipno1), Val(sSLipno2), saItemCd(iL))
                End If
            End If
            
            If iCount > iPrPageCnt Then
                Printer.NewPage
            End If
            
        Next
    End If
    
    Printer.EndDoc
    iPrPageCnt = iPrPageCnt + 1
    Return


BanZul_Sub:
    Printer.FontSize = 5
    Printer.Print ""
    Printer.FontSize = 9
    Return



Get_RefData:
    Dim adoRef      As ADODB.Recordset
    
    sMin = "": sMax = ""
    
    strSql = ""
    strSql = strSql & " SELECT * "
    strSql = strSql & " FROM   TWEXAM_REFDATA"
    strSql = strSql & " WHERE  ITEMCODE  = '" & sRitemCd & "'"
    strSql = strSql & " AND    AGEMIN   <=  " & Val(sAge)
    strSql = strSql & " AND    AGEMAX   >=  " & Val(sAge)
    strSql = strSql & " AND    APPDATE   =     (SELECT MAX(APPDATE)"
    strSql = strSql & "                         FROM   TWEXAM_REFDATA"
    strSql = strSql & "                         WHERE  ITEMCODE = '" & sRitemCd & "'"
    strSql = strSql & "                         AND    AGEMIN  <=  " & Val(sAge)
    strSql = strSql & "                         AND    AGEMAX  >=  " & Val(sAge) & ")"
    
    If adoSetOpen(strSql, adoRef) Then
        If sSex = "M" Then
            RSet sMin = adoRef.Fields("M_MIN").Value & ""
            sMax = adoRef.Fields("M_MAX").Value & "": End If
        If sSex = "F" Then
            RSet sMin = adoRef.Fields("F_MIN").Value & ""
            sMax = adoRef.Fields("F_MAX").Value & "": End If
        Call adoSetClose(adoRef)
    End If
    
    Return


Print_Text_Ret:
    Dim sRetCham    As String
    Dim sRet        As String
    Dim adoCham     As ADODB.Recordset
    
    strSql = ""
    strSql = strSql & " SELECT TO_CHAR(a.JeobsuDt,'YYYY-MM-DD') JeobsuDt,"
    strSql = strSql & "        a.ItemCd, a.Result1, b.ItemNM, b.MinCham, b.MaxCham, a.Chamgo"
    strSql = strSql & " FROM   TWEXAM_GENERAL_SUB a,"
    strSql = strSql & "        TWEXAM_ITEMML      b "
    strSql = strSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.Ptno     = '" & sPtno & "'"
    strSql = strSql & " AND    a.SLIPNO1  =  " & Val(sSLipno1)
    strSql = strSql & " AND    a.Verify   = 'Y'"
    strSql = strSql & " AND    a.ItemCd   = b.CodeKy(+)"
    
    If False = adoSetOpen(strSql, adoCham) Then Return
    
    Do Until adoCham.EOF
        sRitemCd = adoCham.Fields("ItemCd").Value & ""
        sRitemNm = adoCham.Fields("ItemNm").Value & ""
        
        sRet = Trim$(adoCham.Fields("Result1").Value & "")
        sRetCham = Trim$(adoCham.Fields("Chamgo").Value & "")
        GoSub Check_Bit_Chamgo
        Printer.Print sRitemCd & " " & sRitemNm
        Printer.Print "결과:__ "
        Printer.Print "        " & sRet
        If Trim$(sRetCham) <> "" Then
            Printer.Print "참고사항: "
            Printer.Print sRetCham
        End If
        
        adoCham.MoveNext
    Loop
    Call adoSetClose(adoCham)
    
    
    Printer.Print ""
    Printer.Print ""
    If Trim$(sRemark) <> "" Then
        Printer.Print "☞ Remark ___________________________"
        Printer.Print ""
        Printer.Print sRemark
    End If
    
    
    Return
    
Check_Bit_Chamgo:
    Dim nLength As Double
    Dim sTarget As String
    Dim nCnt    As Integer
    nLength = Len(sRetCham)
    
    nCnt = 1
    For i = 1 To nLength
        If nCnt > 62 Then
            sTarget = sTarget & vbCrLf & Mid(sRetCham, i, 1)
            nCnt = 1
        Else
            sTarget = sTarget & Mid(sRetCham, i, 1)
            nCnt = nCnt + 1
        End If
        
    Next
    sRetCham = sTarget
    Return

End Sub

Private Sub cmdPrMicro_Click()
    Dim sSLipno1    As String
    Dim sSLipno2    As String
    Dim sPtno       As String * 8
    Dim sJeobsuDt   As String
    Dim sSex        As String
    Dim sGeomsaDt   As String
    Dim sAge        As String
    Dim sRemark     As String
    Dim sSname      As String * 10
    Dim sRitemCd    As String * 8
    Dim sRitemNm    As String * 30
    Dim sDanWi      As String * 6
    Dim sResult     As String * 12
    Dim sResult42   As String
    Dim sMin        As String * 6
    Dim sMax        As String * 6
    Dim sRoomCode   As String
    Dim sDeptName   As String
    Dim sSamplename As String
    Dim iSensCount      As Integer
    Dim iSensOrderName  As String
    
    
    If vbNo = MsgBox(" Print 하시겠습니까?", vbQuestion + vbYesNo, "Printing 할까요?") Then
        Exit Sub
    End If
    
    Printer.Orientation = vbPRORPortrait

    For i = 1 To ssResult.DataRowCnt
        ssResult.Row = i
        ssResult.Col = 1
            
        If ssResult.Value = True Then
            ssResult.Col = 2: sJeobsuDt = ssResult.Text
            ssResult.Col = 3: sSLipno1 = ssResult.Text
            ssResult.Col = 4: sSLipno2 = ssResult.Text
            ssResult.Col = 6: sPtno = ssResult.Text
            GoSub Main_Process
        End If
    Next
    Exit Sub
    
    
Main_Process:
    strSql = ""
    strSql = strSql & " SELECT TO_CHAR(a.JeobsuDt, 'yyyy-MM-dd') JeobsuDt,"
    strSql = strSql & "        a.SLipno1, a.SLipno2, a.Ptno, b.Sname, b.Sex, a.AgeYY,"
    strSql = strSql & "        TO_CHAR(a.GeomsaDt,'yyyy-MM-dd') GeomsaDt,"
    strSql = strSql & "        a.RoomCode, a.GeomsaCM, c.DeptNamek, d.WardCode"
    strSql = strSql & " FROM   TWEXAM_GENERAL  a,"
    strSql = strSql & "        TWEXAM_IDNOMST  b,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT      c,"
    strSql = strSql & "        TW_MIS_PMPA.TWBas_ROOM      d "
    strSql = strSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.Slipno1  =  " & Val(sSLipno1)
    strSql = strSql & " AND    a.SLipno2  =  " & Val(sSLipno2)
    strSql = strSql & " AND    a.PTNO     = '" & sPtno & "'"
    strSql = strSql & " AND    a.Ptno     = b.Ptno(+)"
    strSql = strSql & " AND    a.DeptCode = c.DeptCode(+)"
    strSql = strSql & " AND    a.RoomCode = d.RoomCode(+)"
    
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    sJeobsuDt = adoSet.Fields("JeobsuDt").Value & ""
    sSLipno1 = adoSet.Fields("SLipno1").Value & ""
    sSLipno2 = adoSet.Fields("SLipno2").Value & ""
    sPtno = adoSet.Fields("Ptno").Value & ""
    sSname = adoSet.Fields("Sname").Value & ""
    sSex = adoSet.Fields("Sex").Value & ""
    sGeomsaDt = adoSet.Fields("GeomsaDt").Value & ""
    
    sRoomCode = adoSet.Fields("RoomCode").Value & ""
    If Trim(sRoomCode) <> "" Then
        sRoomCode = adoSet.Fields("WardCode").Value & "/" & sRoomCode
    End If
    
    sDeptName = adoSet.Fields("DeptNamek").Value & ""
    sRemark = adoSet.Fields("GeomsaCM").Value & ""
    sAge = adoSet.Fields("AgeYY").Value & ""
    Call adoSetClose(adoSet)
    
    
    GoSub PrintHead_RTN
    GoSub Print_OK_RTN
    
    Return


PrintHead_RTN:
    Dim sDeptCode   As String * 10
    Dim sAgeYY      As String
    Dim sJDT        As String
    Dim sGDT        As String
    Dim sSlipTitle  As String
    Dim sGeomsaJa   As String
    Dim sLabno      As String * 5
    Dim adoSpec     As ADODB.Recordset
    Dim adoGen      As ADODB.Recordset
    Dim sStat       As String
    
    
    strSql = ""
    strSql = strSql & " SELECT codenm"
    strSql = strSql & " FROM   TWEXAM_SPECODE"
    strSql = strSql & " WHERE  CODEGU = '12'"
    strSql = strSql & " AND    CODEKY = '" & sSLipno1 & "'"
    If adoSetOpen(strSql, adoSpec) Then
        sSlipTitle = adoSpec.Fields("Codenm").Value & ""
        Call adoSetClose(adoSpec)
    End If
    
'o  strSql = ""
'o  strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_DEPT INDEX_DEPT0) */"

    strSql = ""
    strSql = strSql & " SELECT a.RoomCode, a.Sex, a.AgeYY, a.SLipno1, a.SLipno2, a.Status,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDT,'YYYY-MM-DD') JeobsuDt, a.JeobsuT1, a.JeobsuT2,"
    strSql = strSql & "        TO_CHAR(a.GeomsaDT,'YYYY-MM-DD') GeomsaDt, a.GeomsaT1, a.GeomsaT2,"
    strSql = strSql & "        a.GeomsaCM, b.DeptNamek, c.Name, d.Codenm"
    strSql = strSql & " FROM   TWEXAM_GENERAL a,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT     b,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PASS     c,"
    strSql = strSql & "        TWEXAM_Sample  d "
    strSql = strSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.Ptno     = '" & sPtno & "'"
    strSql = strSql & " AND    a.Slipno1  =  " & Val(sSLipno1)
    strSql = strSql & " AND    a.Slipno2  =  " & Val(sSLipno2)
    strSql = strSql & " AND    a.DeptCode = b.DeptCode(+)"
    strSql = strSql & " AND    a.Geomsaja = c.idNumber(+)"
    strSql = strSql & " AND    a.GeomchCd = d.Code(+)"
    
    If False = adoSetOpen(strSql, adoGen) Then Return
    
    sDeptCode = adoGen.Fields("DeptNamek").Value & ""
    sRoomCode = adoGen.Fields("RoomCode").Value & ""
    sSex = adoGen.Fields("Sex").Value & ""
    sAgeYY = adoGen.Fields("AgeYY").Value & ""
    sJDT = adoGen.Fields("JeobsuDt").Value & " " & adoGen.Fields("JeobsuT1").Value & ":" & adoGen.Fields("JeobsuT2").Value & ""
    sGDT = adoGen.Fields("GeomsaDt").Value & " " & adoGen.Fields("GeomsaT1").Value & ":" & adoGen.Fields("GeomsaT2").Value & ""
    sRemark = Trim(adoGen.Fields("GeomsaCm").Value & "")
    sGeomsaJa = Trim(adoGen.Fields("Name").Value & "")
    sSamplename = Trim(adoGen.Fields("Codenm").Value & "")
    sLabno = adoGen.Fields("SLipno2").Value & ""
    
    Select Case Trim(adoGen.Fields("Status").Value & "")
        Case "R": sStat = "접수중"
        Case "P": sStat = "부분결과"
        Case "U": sStat = "미확인"
        Case "C": sStat = "결과완료"
        Case "X": sStat = "이상Data"
        Case Else: sStat = ""
    End Select
    
    Call adoSetClose(adoGen)
    
    Printer.FontName = "바탕체"
    Printer.FontSize = "12"
    Printer.FontBold = True
    Printer.FontItalic = True
    
    Printer.Print sSlipTitle & "       Result Report   " & sLabno & "  (" & Trim(sStat) & ")"
    Printer.ForeColor = RGB(192, 192, 192)
    Printer.Print "━━━━━━━━━━━━━━━━━━━━━━━━━"
    Printer.ForeColor = RGB(0, 0, 0)
    
    Printer.FontName = "바탕체"
    Printer.FontSize = 9
    Printer.FontBold = False
    Printer.FontItalic = False
    
    Printer.Print "등록No: " & sPtno; Tab(40); "검  체: " & sSamplename
    Printer.Print "성  명: " & Trim(sSname) & "[" & sSex & "/" & sAgeYY & "]"; Tab(40); "검사자: " & sGeomsaJa
    Printer.Print "진료과: " & sDeptName; Tab(40); "접수일: " & sJDT
    Printer.Print "병  실: " & sRoomCode; Tab(40); "검사일: " & sGDT

    If Left(sSLipno1, 1) = "4" Then
        Printer.ForeColor = RGB(192, 192, 192)
        Printer.Print "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
        Printer.ForeColor = RGB(0, 0, 0)
        Printer.Print "       검사항목             :         검사결과                      "
        Printer.ForeColor = RGB(192, 192, 192)
        Printer.Print "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
        Printer.ForeColor = RGB(0, 0, 0)
    Else
        Printer.ForeColor = RGB(192, 192, 192)
        Printer.Print "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
        Printer.ForeColor = RGB(0, 0, 0)
        Printer.Print "       검사항목             :  검사결과 :     참고치      :   단위  "
        Printer.ForeColor = RGB(192, 192, 192)
        Printer.Print "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
        Printer.ForeColor = RGB(0, 0, 0)
    End If
    Return
    
    
    
Print_OK_RTN:
    Dim j               As Integer
    Dim adoGsub         As ADODB.Recordset
    Dim sSensFlag       As String * 1
    Dim nPrcnt          As Integer
    Dim saItemCd(10)    As String
    Dim iL              As Integer
    Dim sFlagAbnormal   As String
    Dim iRcodeCount     As Integer
    Dim iResultCount    As Integer
    
    sSensFlag = ""
    iRcodeCount = 0
    iResultCount = 0
    
    strSql = ""
    strSql = strSql & " SELECT TO_CHAR(a.JeobsuDt,'YYYY-MM-DD') JeobsuDt,"
    strSql = strSql & "        a.ItemCd, a.Result1, a.Result2, a.Result3, a.Result4, a.Result5, "
    strSql = strSql & "        b.ItemNM, b.MinCham, b.MaxCham, b.DanWi, b.MinDanger, b.MaxDanger,"
    strSql = strSql & "        b.ResultW"
    strSql = strSql & " FROM   TWEXAM_GENERAL_SUB a,"
    strSql = strSql & "        TWEXAM_ITEMML      b "
    strSql = strSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.Ptno     = '" & sPtno & "'"
    strSql = strSql & " AND    a.Slipno1  =  " & Val(sSLipno1)
    strSql = strSql & " AND    a.Slipno2  =  " & Val(sSLipno2)
    'strSql = strSql & " AND    a.Verify   = 'Y'"
    strSql = strSql & " AND    a.ItemCd   = b.CodeKy(+)"
    strSql = strSql & " ORDER  BY a.itemCd"
    
    If False = adoSetOpen(strSql, adoGsub) Then Return
    
    
    Printer.FontName = "바탕체"
    Printer.FontSize = 9
    Printer.FontBold = False
    Printer.FontItalic = False
    
    Do Until adoGsub.EOF
        sRitemCd = adoGsub.Fields("ItemCd").Value & ""
        sRitemNm = adoGsub.Fields("ItemNm").Value & ""
                
        If Trim(adoGsub.Fields("ResultW").Value & "") = "S" Then
            Printer.FontSize = 5: Printer.Print ""
            Printer.FontName = "바탕체": Printer.FontSize = 12:  Printer.FontBold = True: Printer.FontItalic = False
            sResult42 = "다음장참조"
            Printer.Print "  " & sRitemNm & sResult42
            
        Else
            Printer.FontName = "바탕체": Printer.FontSize = 10:  Printer.FontBold = False: Printer.FontItalic = False
            sResult42 = adoGsub.Fields("Result1").Value & ""
            Printer.Print "  " & sRitemNm & Trim(sResult42)
        End If
        adoGsub.MoveNext
    Loop
    
    Printer.FontName = "바탕체": Printer.FontSize = 10:  Printer.FontBold = False: Printer.FontItalic = False
    Printer.Print ""
    Printer.Print ""
    Printer.Print "☞ Remark ___________________________"
    Printer.Print ""
    Printer.Print " :" & sRemark
    
    Printer.NewPage
    
    adoGsub.MoveFirst
    Do Until adoGsub.EOF
        If Trim(adoGsub.Fields("ResultW").Value & "") = "S" Then
            GoSub Print_Sens_Loop
        End If
        adoGsub.MoveNext
    Loop
    Call adoSetClose(adoGsub)
    Printer.EndDoc
    Return
    

Print_Sens_Loop:
    Dim sTmpRcode       As String
    Dim sDispRname      As String * 25
    
    sRitemCd = adoGsub.Fields("ItemCd").Value & ""
    sRitemNm = adoGsub.Fields("ItemNm").Value & ""
    RSet sDanWi = adoGsub.Fields("DanWi").Value & ""
    LSet sResult = adoGsub.Fields("Result1").Value & ""
    
    
    Printer.FontName = "바탕체"
    Printer.FontSize = "12"
    Printer.FontBold = True
    Printer.FontItalic = True

    Printer.Print sSlipTitle & "(SENSTIVITY)  Result Report   " & sLabno & "(" & Trim(sStat) & ")"
    Printer.FontName = "바탕체"
    Printer.FontSize = 9
    Printer.FontBold = False
    Printer.FontItalic = False
    
    Printer.ForeColor = RGB(192, 192, 192)
    Printer.Print "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    Printer.ForeColor = RGB(0, 0, 0)
    Printer.Print "등록No: " & sPtno; Tab(40); "검  체: " & sSamplename
    Printer.Print "성  명: " & Trim(sSname) & "[" & sSex & "/" & sAgeYY & "]"; Tab(40); "검사자: " & sGeomsaJa
    Printer.Print "진료과: " & sDeptName; Tab(40); "접수일: " & sJDT
    Printer.Print "병  실: " & sRoomCode; Tab(40); "검사일: " & sGDT
    Printer.ForeColor = RGB(192, 192, 192)
    Printer.Print "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    Printer.ForeColor = RGB(0, 0, 0)
    
    Printer.FontName = "바탕체": Printer.FontSize = 10:  Printer.FontBold = True: Printer.FontItalic = False
    Printer.Print "Order    : " & sRitemNm
    Printer.Print "보고형태 : " & GET_General_Status(sJeobsuDt, Val(sSLipno1), Val(sSLipno2))
    
    
    iSensCount = Printing_Sens(sJeobsuDt, Val(sSLipno1), Val(sSLipno2), sRitemCd)
                
    If iSensCount > 0 Then
        For nPrcnt = 0 To iSensCount - 1
            If sTmpRcode = SensResult.Rcode(nPrcnt) Then
                Printer.Print "     " & SensResult.AntiName(nPrcnt) & "," & SensResult.Sens(nPrcnt)
                              
            Else
                Printer.Print ""
                Printer.FontName = "바탕체": Printer.FontSize = 9:  Printer.FontBold = True: Printer.FontItalic = False
                Printer.Print "@." & _
                              SensResult.Rcode(nPrcnt) & vbCrLf & _
                              "     " & SensResult.AntiName(nPrcnt) & "," & SensResult.Sens(nPrcnt)
            End If
            sTmpRcode = SensResult.Rcode(nPrcnt)
        Next
    Else
        If Trim(sRitemCd) <> "" Then
            Printer.Print ""
            Printer.FontName = "바탕체": Printer.FontSize = 9:  Printer.FontBold = True: Printer.FontItalic = False
            Printer.Print Printing_RCode(sJeobsuDt, Val(sSLipno1), Val(sSLipno2), sRitemCd)
        End If
    End If
    
    Printer.NewPage
    
    Return




BanZul_Sub:
    Printer.FontSize = 5
    Printer.Print ""
    Printer.FontSize = 9
    Return


Print_Text_Ret:
    Dim sRetCham    As String
    Dim sRet        As String
    Dim adoCham     As ADODB.Recordset
    
    strSql = ""
    strSql = strSql & " SELECT TO_CHAR(a.JeobsuDt,'YYYY-MM-DD') JeobsuDt,"
    strSql = strSql & "        a.ItemCd, a.Result1, b.ItemNM, b.MinCham, b.MaxCham, a.Chamgo"
    strSql = strSql & " FROM   TWEXAM_GENERAL_SUB a,"
    strSql = strSql & "        TWEXAM_ITEMML      b "
    strSql = strSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.Ptno     = '" & sPtno & "'"
    strSql = strSql & " AND    a.SLIPNO1  =  " & Val(sSLipno1)
    strSql = strSql & " AND    a.Verify   = 'Y'"
    strSql = strSql & " AND    a.ItemCd   = b.CodeKy(+)"
    
    If False = adoSetOpen(strSql, adoCham) Then Return
    
    Do Until adoCham.EOF
        sRitemCd = adoCham.Fields("ItemCd").Value & ""
        sRitemNm = adoCham.Fields("ItemNm").Value & ""
        
        sRet = Trim$(adoCham.Fields("Result1").Value & "")
        sRetCham = Trim$(adoCham.Fields("Chamgo").Value & "")
        GoSub Check_Bit_Chamgo
        Printer.Print sRitemCd & " " & sRitemNm
        Printer.Print "결과:__ "
        Printer.Print "        " & sRet
        If Trim$(sRetCham) <> "" Then
            Printer.Print "참고사항: "
            Printer.Print sRetCham
        End If
        
        adoCham.MoveNext
    Loop
    Call adoSetClose(adoCham)
    
    
    Printer.Print ""
    Printer.Print ""
    If Trim$(sRemark) <> "" Then
        Printer.Print "☞ Remark ___________________________"
        Printer.Print ""
        Printer.Print sRemark
    End If
    
    
    Return
    
Check_Bit_Chamgo:
    Dim nLength As Double
    Dim sTarget As String
    Dim nCnt    As Integer
    nLength = Len(sRetCham)
    
    nCnt = 1
    For i = 1 To nLength
        If nCnt > 62 Then
            sTarget = sTarget & vbCrLf & Mid(sRetCham, i, 1)
            nCnt = 1
        Else
            sTarget = sTarget & Mid(sRetCham, i, 1)
            nCnt = nCnt + 1
        End If
        
    Next
    sRetCham = sTarget
    Return

End Sub

Private Sub cmdQuery_Click()
    Dim sFrDate         As String
    Dim sToDate         As String
    Dim sWhere          As String * 1
    
    
    Call SpreadSetClear(ssResult)
    
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd") & " " & txtFrJeobsuT1.Text & ":" & txtFrJeobsuT2.Text
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd") & " " & txtToJeobsuT1.Text & ":" & txtToJeobsuT2.Text
    
    If optWhere(0).Value = True Then sWhere = "R"
    If optWhere(1).Value = True Then sWhere = "P"
    If optWhere(2).Value = True Then sWhere = "C'"
    
'o  strSql = ""
'o  strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INDEX_PATIENT0) */"
   
    strSql = ""
    strSql = strSql & " SELECT a.*, a.RowID,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDT, 'YYYY-MM-DD') JeobsuDt, "
    strSql = strSql & "        TO_CHAR(a.OrderDt,  'YYYY-MM-DD') OrderDt,  "
    strSql = strSql & "        TO_CHAR(a.GeomsaDt, 'YYYY-MM-DD') GeomsaDt, "
    strSql = strSql & "        b.Sname,    b.AgeYY,  b.Sex,  c.Deptnamek, d.Drname, e.Name, f.Codenm"
    strSql = strSql & " FROM   TWEXAM_GENERAL  a, "
    strSql = strSql & "        TWEXAM_IDNOMST  b, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT      c, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR    d, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PASS      e, "
    strSql = strSql & "        TWEXAM_Sample   f  "
    

    If Option1.Value = True Then '접수일자
        strSql = strSql & " WHERE  LTRIM(TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD')) || ' ' || "
        strSql = strSql & "        LTRIM(TO_CHAR(a.JeobsuT1, '00')) || ':' || "
        strSql = strSql & "        LTRIM(TO_CHAR(a.JeobsuT2, '00'))   BETWEEN  '" & sFrDate & "'"
        strSql = strSql & "                                           AND      '" & sToDate & "'"
    End If
    
    If Option2.Value = True Then '검사일자
        strSql = strSql & " WHERE  LTRIM(TO_CHAR(a.GeomsaDt, 'YYYY-MM-DD')) || ' ' || "
        strSql = strSql & "        LTRIM(TO_CHAR(a.GeomsaT1, '00')) || ':' || "
        strSql = strSql & "        LTRIM(TO_CHAR(a.GeomsaT2, '00'))   BETWEEN  '" & sFrDate & "'"
        strSql = strSql & "                                           AND      '" & sToDate & "'"
    End If
    
    strSql = strSql & " AND    a.Slipno1   = " & Val(Left(cmbSLip.Text, 2))
'    strSql = strSql & " AND    a.GbCh      = 'Y'"
    strSql = strSql & " AND    a.DeptCode  != 'GE'"            '일반검진 제외
    strSql = strSql & " AND    a.Status    = '" & sWhere & "'"
    strSql = strSql & " AND    a.Ptno      = b.Ptno(+)"
    strSql = strSql & " AND    a.DeptCode  = c.DeptCode(+)"
    strSql = strSql & " AND    a.DrCode    = d.DrCode(+)"
    strSql = strSql & " AND    a.GeomchCD  = f.Code(+)"
    strSql = strSql & " AND    a.Geomsaja  = e.Idnumber(+)"
    strSql = strSql & " AND   (e.Programid = ' ' OR e.Programid IS NULL)"
    strSql = strSql & " ORDER  BY a.Jeobsudt, a.SLipno2"
    
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        ssResult.Row = ssResult.DataRowCnt + 1
        ssResult.Col = 1:
        ssResult.Col = 2:  ssResult.Text = adoSet.Fields("JeobsuDt").Value & ""
        ssResult.Col = 3:  ssResult.Text = adoSet.Fields("SLipno1").Value & ""
        ssResult.Col = 4:  ssResult.Text = adoSet.Fields("SLipno2").Value & ""
        ssResult.Col = 5:  ssResult.Text = adoSet.Fields("RoomCode").Value & ""
        ssResult.Col = 6:  ssResult.Text = adoSet.Fields("Ptno").Value & ""
        ssResult.Col = 7:  ssResult.Text = adoSet.Fields("Sname").Value & ""
        ssResult.Col = 8:  ssResult.Text = adoSet.Fields("Sex").Value & ""
        ssResult.Col = 9:  ssResult.Text = adoSet.Fields("AgeYY").Value & ""
        ssResult.Col = 10: ssResult.Text = adoSet.Fields("GeomsaDt").Value & ""
        ssResult.Col = 11: ssResult.Text = adoSet.Fields("GeomsaT1").Value & ":" & _
                                           adoSet.Fields("GeomsaT2").Value & ""
        ssResult.Col = 12: ssResult.Text = adoSet.Fields("Geomsaja").Value & ""
        ssResult.Col = 13: ssResult.Text = adoSet.Fields("GeomchCd").Value & ""
        ssResult.Col = 14: ssResult.Text = adoSet.Fields("DeptCode").Value & ""
        ssResult.Col = 15: ssResult.Text = adoSet.Fields("DeptNameK").Value & ""
        ssResult.Col = 16: ssResult.Text = adoSet.Fields("GeomsaCM").Value & ""
        ssResult.Col = 17: ssResult.Text = adoSet.Fields("JeobsuT1").Value & ":" & _
                                          adoSet.Fields("JeobsuT2").Value & ""
        ssResult.Col = 18: ssResult.Text = adoSet.Fields("OrderDt").Value & "'"
        
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    

End Sub

Private Sub cmdQuery1_Click()
    Dim sFrDate         As String
    Dim sToDate         As String
    Dim sWhere          As String * 1
    
    
    Call SpreadSetClear(ssResult)
    
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    If optWhere(0).Value = True Then sWhere = "R"
    If optWhere(1).Value = True Then sWhere = "P"
    If optWhere(2).Value = True Then sWhere = "C'"
    
'o  strSql = ""
'o  strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INDEX_PATIENT0) */"

    strSql = ""
    strSql = strSql & " SELECT a.*, a.RowID,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDT, 'YYYY-MM-DD') JeobsuDt, "
    strSql = strSql & "        TO_CHAR(a.OrderDt,  'YYYY-MM-DD') OrderDt,  "
    strSql = strSql & "        TO_CHAR(a.GeomsaDt, 'YYYY-MM-DD') GeomsaDt, "
    strSql = strSql & "        b.Sname,     b.AgeYY, b.Sex,  c.Deptnamek, d.Drname, e.Name, f.Codenm"
    strSql = strSql & " FROM   TWEXAM_GENERAL  a, "
    strSql = strSql & "        TWEXAM_IDNOMST  b, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT      c, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR    d, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PASS      e, "
    strSql = strSql & "        TWEXAM_Sample   f  "
    strSql = strSql & " WHERE  a.JeobsuDt >= TO_DATE('" & sFrDate & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.JeobsuDt <= TO_DATE('" & sToDate & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.Slipno1   = " & Val(Left(cmbSLip.Text, 2))
    strSql = strSql & " AND    a.DeptCode  != 'GE'"
    strSql = strSql & " AND    a.Ptno      = '" & txtPtno.Text & "'"
    strSql = strSql & " AND    a.GbCh      = 'Y'"
    strSql = strSql & " AND    a.Status    = '" & sWhere & "'"
    strSql = strSql & " AND    a.Ptno      = b.Ptno(+)"
    strSql = strSql & " AND    a.DeptCode  = c.DeptCode(+)"
    strSql = strSql & " AND    a.DrCode    = d.DrCode(+)"
    strSql = strSql & " AND    a.GeomchCD  = f.Code(+)"
    strSql = strSql & " AND    a.Geomsaja  = e.Idnumber(+)"
    strSql = strSql & " AND   (e.Programid = ' ' OR e.Programid IS NULL)"
    strSql = strSql & " ORDER  BY a.Jeobsudt, a.SLipno2"
    
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        ssResult.Row = ssResult.DataRowCnt + 1
        ssResult.Col = 1:
        ssResult.Col = 2:  ssResult.Text = adoSet.Fields("JeobsuDt").Value & ""
        ssResult.Col = 3:  ssResult.Text = adoSet.Fields("SLipno1").Value & ""
        ssResult.Col = 4:  ssResult.Text = adoSet.Fields("SLipno2").Value & ""
        ssResult.Col = 5:  ssResult.Text = adoSet.Fields("RoomCode").Value & ""
        ssResult.Col = 6:  ssResult.Text = adoSet.Fields("Ptno").Value & ""
        ssResult.Col = 7:  ssResult.Text = adoSet.Fields("Sname").Value & ""
        ssResult.Col = 8:  ssResult.Text = adoSet.Fields("Sex").Value & ""
        ssResult.Col = 9:  ssResult.Text = adoSet.Fields("AgeYY").Value & ""
        ssResult.Col = 10: ssResult.Text = adoSet.Fields("GeomsaDt").Value & ""
        ssResult.Col = 11: ssResult.Text = adoSet.Fields("GeomsaT1").Value & ":" & _
                                          adoSet.Fields("GeomsaT2").Value & ""
        ssResult.Col = 12: ssResult.Text = adoSet.Fields("Geomsaja").Value & ""
        ssResult.Col = 13: ssResult.Text = adoSet.Fields("GeomchCd").Value & ""
        ssResult.Col = 14: ssResult.Text = adoSet.Fields("DeptCode").Value & ""
        ssResult.Col = 15: ssResult.Text = adoSet.Fields("DeptNameK").Value & ""
        ssResult.Col = 16: ssResult.Text = adoSet.Fields("GeomsaCM").Value & ""
        ssResult.Col = 17: ssResult.Text = adoSet.Fields("JeobsuT1").Value & ":" & _
                                          adoSet.Fields("JeobsuT2").Value & ""
        ssResult.Col = 18: ssResult.Text = adoSet.Fields("OrderDt").Value & "'"
        
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)

End Sub

Private Sub cmdQuery2_Click()
    Dim sFrDate         As String
    Dim sToDate         As String
    Dim sWhere          As String * 1
    
    
    Call SpreadSetClear(ssResult)
    
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    If optWhere(0).Value = True Then sWhere = "R"
    If optWhere(1).Value = True Then sWhere = "P"
    If optWhere(2).Value = True Then sWhere = "C'"
    
'o  strSql = ""
'o  strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INDEX_PATIENT0) */"

    strSql = ""
    strSql = strSql & " SELECT a.*, a.RowID,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDT, 'YYYY-MM-DD') JeobsuDt, "
    strSql = strSql & "        TO_CHAR(a.OrderDt,  'YYYY-MM-DD') OrderDt,  "
    strSql = strSql & "        TO_CHAR(a.GeomsaDt, 'YYYY-MM-DD') GeomsaDt, "
    strSql = strSql & "        b.Sname,      b.AgeYY,  b.Sex,  c.Deptnamek, d.Drname, e.Name, f.Codenm"
    strSql = strSql & " FROM   TWEXAM_GENERAL  a, "
    strSql = strSql & "        TWEXAM_IDNOMST  b, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT      c, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR    d, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PASS      e, "
    strSql = strSql & "        TWEXAM_Sample   f  "
    strSql = strSql & " WHERE  a.JeobsuDt >= TO_DATE('" & sFrDate & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.JeobsuDt <= TO_DATE('" & sToDate & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.Slipno1   = " & Val(Left(cmbSLip.Text, 2))
    strSql = strSql & " AND    b.Sname     LIKE '" & txtSname.Text & "%'"
    strSql = strSql & " AND    a.GbCh      = 'Y'"
    strSql = strSql & " AND    a.DeptCode  != 'GE'"
    strSql = strSql & " AND    a.Status    = '" & sWhere & "'"
    strSql = strSql & " AND    a.Ptno      = b.Ptno(+)"
    strSql = strSql & " AND    a.DeptCode  = c.DeptCode(+)"
    strSql = strSql & " AND    a.DrCode    = d.DrCode(+)"
    strSql = strSql & " AND    a.GeomchCD  = f.Code(+)"
    strSql = strSql & " AND    a.Geomsaja  = e.Idnumber(+)"
    strSql = strSql & " AND   (e.Programid = ' ' OR e.Programid IS NULL)"
    strSql = strSql & " ORDER  BY a.Jeobsudt, a.SLipno2"
    
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        ssResult.Row = ssResult.DataRowCnt + 1
        ssResult.Col = 1:
        ssResult.Col = 2:  ssResult.Text = adoSet.Fields("JeobsuDt").Value & ""
        ssResult.Col = 3:  ssResult.Text = adoSet.Fields("SLipno1").Value & ""
        ssResult.Col = 4:  ssResult.Text = adoSet.Fields("SLipno2").Value & ""
        ssResult.Col = 5:  ssResult.Text = adoSet.Fields("RoomCode").Value & ""
        ssResult.Col = 6:  ssResult.Text = adoSet.Fields("Ptno").Value & ""
        ssResult.Col = 7:  ssResult.Text = adoSet.Fields("Sname").Value & ""
        ssResult.Col = 8:  ssResult.Text = adoSet.Fields("Sex").Value & ""
        ssResult.Col = 9:  ssResult.Text = adoSet.Fields("AgeYY").Value & ""
        ssResult.Col = 10: ssResult.Text = adoSet.Fields("GeomsaDt").Value & ""
        ssResult.Col = 11: ssResult.Text = adoSet.Fields("GeomsaT1").Value & ":" & _
                                          adoSet.Fields("GeomsaT2").Value & ""
        ssResult.Col = 12: ssResult.Text = adoSet.Fields("Geomsaja").Value & ""
        ssResult.Col = 13: ssResult.Text = adoSet.Fields("GeomchCd").Value & ""
        ssResult.Col = 14: ssResult.Text = adoSet.Fields("DeptCode").Value & ""
        ssResult.Col = 15: ssResult.Text = adoSet.Fields("DeptNameK").Value & ""
        ssResult.Col = 16: ssResult.Text = adoSet.Fields("GeomsaCM").Value & ""
        ssResult.Col = 17: ssResult.Text = adoSet.Fields("JeobsuT1").Value & ":" & _
                                          adoSet.Fields("JeobsuT2").Value & ""
        ssResult.Col = 18: ssResult.Text = adoSet.Fields("OrderDt").Value & "'"
        
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)

End Sub

Private Sub cmdQuery3_Click()
    Dim sFrDate         As String
    Dim sToDate         As String
    Dim sWhere          As String * 1
    
    
    Call SpreadSetClear(ssResult)
    
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    If optWhere(0).Value = True Then sWhere = "R"
    If optWhere(1).Value = True Then sWhere = "P"
    If optWhere(2).Value = True Then sWhere = "C'"
    
'o  strSql = ""
'o  strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INDEX_PATIENT0) */"

    strSql = ""
    strSql = strSql & " SELECT a.*, a.RowID,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDT, 'YYYY-MM-DD') JeobsuDt, "
    strSql = strSql & "        TO_CHAR(a.OrderDt,  'YYYY-MM-DD') OrderDt,  "
    strSql = strSql & "        TO_CHAR(a.GeomsaDt, 'YYYY-MM-DD') GeomsaDt, "
    strSql = strSql & "        b.Sname,       b.AgeYY, b.Sex,  c.Deptnamek, d.Drname, e.Name, f.Codenm"
    strSql = strSql & " FROM   TWEXAM_GENERAL  a, "
    strSql = strSql & "        TWEXAM_IDNOMST  b, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT      c, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR    d, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PASS      e, "
    strSql = strSql & "        TWEXAM_Sample   f  "
    strSql = strSql & " WHERE  a.JeobsuDt >= TO_DATE('" & sFrDate & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.JeobsuDt <= TO_DATE('" & sToDate & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.Slipno1   = " & Val(Left(cmbSLip.Text, 2))
    strSql = strSql & " AND    a.SLipno2  >= " & Val(txtLabno.Text)
    strSql = strSql & " AND    a.DeptCode  != 'GE'"
    strSql = strSql & " AND    a.SLipno2  <= " & Val(txtToLabno.Text)
    strSql = strSql & " AND    a.GbCh      = 'Y'"
    strSql = strSql & " AND    a.Status    = '" & sWhere & "'"
    strSql = strSql & " AND    a.Ptno      = b.Ptno(+)"
    strSql = strSql & " AND    a.DeptCode  = c.DeptCode(+)"
    strSql = strSql & " AND    a.DrCode    = d.DrCode(+)"
    strSql = strSql & " AND    a.GeomchCD  = f.Code(+)"
    strSql = strSql & " AND    a.Geomsaja  = e.Idnumber(+)"
    strSql = strSql & " AND   (e.Programid = ' ' OR e.Programid IS NULL)"
    strSql = strSql & " ORDER  BY a.Jeobsudt, a.SLipno2"
    
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        ssResult.Row = ssResult.DataRowCnt + 1
        ssResult.Col = 1:
        ssResult.Col = 2:  ssResult.Text = adoSet.Fields("JeobsuDt").Value & ""
        ssResult.Col = 3:  ssResult.Text = adoSet.Fields("SLipno1").Value & ""
        ssResult.Col = 4:  ssResult.Text = adoSet.Fields("SLipno2").Value & ""
        ssResult.Col = 5:  ssResult.Text = adoSet.Fields("RoomCode").Value & ""
        ssResult.Col = 6:  ssResult.Text = adoSet.Fields("Ptno").Value & ""
        ssResult.Col = 7:  ssResult.Text = adoSet.Fields("Sname").Value & ""
        ssResult.Col = 8:  ssResult.Text = adoSet.Fields("Sex").Value & ""
        ssResult.Col = 9:  ssResult.Text = adoSet.Fields("AgeYY").Value & ""
        ssResult.Col = 10: ssResult.Text = adoSet.Fields("GeomsaDt").Value & ""
        ssResult.Col = 11: ssResult.Text = adoSet.Fields("GeomsaT1").Value & ":" & _
                                          adoSet.Fields("GeomsaT2").Value & ""
        ssResult.Col = 12: ssResult.Text = adoSet.Fields("Geomsaja").Value & ""
        ssResult.Col = 13: ssResult.Text = adoSet.Fields("GeomchCd").Value & ""
        ssResult.Col = 14: ssResult.Text = adoSet.Fields("DeptCode").Value & ""
        ssResult.Col = 15: ssResult.Text = adoSet.Fields("DeptNameK").Value & ""
        ssResult.Col = 16: ssResult.Text = adoSet.Fields("GeomsaCM").Value & ""
        ssResult.Col = 17: ssResult.Text = adoSet.Fields("JeobsuT1").Value & ":" & _
                                          adoSet.Fields("JeobsuT2").Value & ""
        ssResult.Col = 18: ssResult.Text = adoSet.Fields("OrderDt").Value & "'"
        
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)

End Sub

Private Sub cmdQuery4_Click()
    Dim sFrDate         As String
    Dim sToDate         As String
    Dim sWhere          As String * 1
    
    
    Call SpreadSetClear(ssResult)
    
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    If optWhere(0).Value = True Then sWhere = "R"
    If optWhere(1).Value = True Then sWhere = "P"
    If optWhere(2).Value = True Then sWhere = "C'"
    
'o  strSql = ""
'o  strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INDEX_PATIENT0) */"

    strSql = ""
    strSql = strSql & " SELECT a.*, a.RowID,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDT, 'YYYY-MM-DD') JeobsuDt, "
    strSql = strSql & "        TO_CHAR(a.OrderDt,  'YYYY-MM-DD') OrderDt,  "
    strSql = strSql & "        TO_CHAR(a.GeomsaDt, 'YYYY-MM-DD') GeomsaDt, "
    strSql = strSql & "        b.Sname,       b.AgeYY, b.Sex,  c.Deptnamek, d.Drname, e.Name, f.Codenm"
    strSql = strSql & " FROM   TWEXAM_GENERAL  a, "
    strSql = strSql & "        TWEXAM_IDNOMST  b, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT      c, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR    d, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PASS      e, "
    strSql = strSql & "        TWEXAM_Sample   f  "
    strSql = strSql & " WHERE  a.JeobsuDt >= TO_DATE('" & sFrDate & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.JeobsuDt <= TO_DATE('" & sToDate & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.Slipno1   = " & Val(Left(cmbSLip.Text, 2))
    strSql = strSql & " AND    a.DeptCode  = '" & Left(cmbDept.Text, 4) & "'"
    strSql = strSql & " AND    a.GbCh      = 'Y'"
    strSql = strSql & " AND    a.DeptCode  != 'GE'"
    strSql = strSql & " AND    a.Status    = '" & sWhere & "'"
    strSql = strSql & " AND    a.Ptno      = b.Ptno(+)"
    strSql = strSql & " AND    a.DeptCode  = c.DeptCode(+)"
    strSql = strSql & " AND    a.DrCode    = d.DrCode(+)"
    strSql = strSql & " AND    a.GeomchCD  = f.Code(+)"
    strSql = strSql & " AND    a.Geomsaja  = e.Idnumber(+)"
    strSql = strSql & " AND   (e.Programid = ' ' OR e.Programid IS NULL)"
    strSql = strSql & " ORDER  BY a.Jeobsudt, a.SLipno2"
    
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        ssResult.Row = ssResult.DataRowCnt + 1
        ssResult.Col = 1:
        ssResult.Col = 2:  ssResult.Text = adoSet.Fields("JeobsuDt").Value & ""
        ssResult.Col = 3:  ssResult.Text = adoSet.Fields("SLipno1").Value & ""
        ssResult.Col = 4:  ssResult.Text = adoSet.Fields("SLipno2").Value & ""
        ssResult.Col = 5:  ssResult.Text = adoSet.Fields("RoomCode").Value & ""
        ssResult.Col = 6:  ssResult.Text = adoSet.Fields("Ptno").Value & ""
        ssResult.Col = 7:  ssResult.Text = adoSet.Fields("Sname").Value & ""
        ssResult.Col = 8:  ssResult.Text = adoSet.Fields("Sex").Value & ""
        ssResult.Col = 9:  ssResult.Text = adoSet.Fields("AgeYY").Value & ""
        ssResult.Col = 10: ssResult.Text = adoSet.Fields("GeomsaDt").Value & ""
        ssResult.Col = 11: ssResult.Text = adoSet.Fields("GeomsaT1").Value & ":" & _
                                          adoSet.Fields("GeomsaT2").Value & ""
        ssResult.Col = 12: ssResult.Text = adoSet.Fields("Geomsaja").Value & ""
        ssResult.Col = 13: ssResult.Text = adoSet.Fields("GeomchCd").Value & ""
        ssResult.Col = 14: ssResult.Text = adoSet.Fields("DeptCode").Value & ""
        ssResult.Col = 15: ssResult.Text = adoSet.Fields("DeptNameK").Value & ""
        ssResult.Col = 16: ssResult.Text = adoSet.Fields("GeomsaCM").Value & ""
        ssResult.Col = 17: ssResult.Text = adoSet.Fields("JeobsuT1").Value & ":" & _
                                          adoSet.Fields("JeobsuT2").Value & ""
        ssResult.Col = 18: ssResult.Text = adoSet.Fields("OrderDt").Value & "'"
        
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)

End Sub

Private Sub cmdQuery5_Click()
    Dim sFrDate         As String
    Dim sToDate         As String
    Dim sWhere          As String * 1
    
    
    Call SpreadSetClear(ssResult)
    
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    If optWhere(0).Value = True Then sWhere = "R"
    If optWhere(1).Value = True Then sWhere = "P"
    If optWhere(2).Value = True Then sWhere = "C'"
    
'o  strSql = ""
'o  strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INDEX_PATIENT0) */"

    strSql = ""
    strSql = strSql & " SELECT a.*, a.RowID,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDT, 'YYYY-MM-DD') JeobsuDt, "
    strSql = strSql & "        TO_CHAR(a.OrderDt,  'YYYY-MM-DD') OrderDt,  "
    strSql = strSql & "        TO_CHAR(a.GeomsaDt, 'YYYY-MM-DD') GeomsaDt, "
    strSql = strSql & "        b.Sname,       b.AgeYY, b.Sex,  c.Deptnamek, d.Drname, e.Name, f.Codenm"
    strSql = strSql & " FROM   TWEXAM_GENERAL  a, "
    strSql = strSql & "        TWEXAM_IDNOMST  b, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT      c, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR    d, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PASS      e, "
    strSql = strSql & "        TWEXAM_Sample   f, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_Room      g  "
    strSql = strSql & " WHERE  a.JeobsuDt >= TO_DATE('" & sFrDate & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.JeobsuDt <= TO_DATE('" & sToDate & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.Slipno1   = " & Val(Left(cmbSLip.Text, 2))
    strSql = strSql & " AND    a.GBIO      = 'I'"
    strSql = strSql & " AND    g.WardCode  = '" & Left(cmbWard.Text, 4) & "'"
    strSql = strSql & " AND    a.GbCh      = 'Y'"
    strSql = strSql & " AND    a.DeptCode  != 'GE'"
    strSql = strSql & " AND    a.Status    = '" & sWhere & "'"
    strSql = strSql & " AND    a.Ptno      = b.Ptno(+)"
    strSql = strSql & " AND    a.DeptCode  = c.DeptCode(+)"
    strSql = strSql & " AND    a.DrCode    = d.DrCode(+)"
    strSql = strSql & " AND    a.GeomchCD  = f.Code(+)"
    strSql = strSql & " AND    a.RoomCode  = g.RoomCode(+)"
    strSql = strSql & " AND    a.Geomsaja  = e.Idnumber(+)"
    strSql = strSql & " AND   (e.Programid = ' ' OR e.Programid IS NULL)"
    strSql = strSql & " ORDER  BY a.Jeobsudt, a.SLipno2"
    
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        ssResult.Row = ssResult.DataRowCnt + 1
        ssResult.Col = 1:
        ssResult.Col = 2:  ssResult.Text = adoSet.Fields("JeobsuDt").Value & ""
        ssResult.Col = 3:  ssResult.Text = adoSet.Fields("SLipno1").Value & ""
        ssResult.Col = 4:  ssResult.Text = adoSet.Fields("SLipno2").Value & ""
        ssResult.Col = 5:  ssResult.Text = adoSet.Fields("RoomCode").Value & ""
        ssResult.Col = 6:  ssResult.Text = adoSet.Fields("Ptno").Value & ""
        ssResult.Col = 7:  ssResult.Text = adoSet.Fields("Sname").Value & ""
        ssResult.Col = 8:  ssResult.Text = adoSet.Fields("Sex").Value & ""
        ssResult.Col = 9:  ssResult.Text = adoSet.Fields("AgeYY").Value & ""
        ssResult.Col = 10: ssResult.Text = adoSet.Fields("GeomsaDt").Value & ""
        ssResult.Col = 11: ssResult.Text = adoSet.Fields("GeomsaT1").Value & ":" & _
                                          adoSet.Fields("GeomsaT2").Value & ""
        ssResult.Col = 12: ssResult.Text = adoSet.Fields("Geomsaja").Value & ""
        ssResult.Col = 13: ssResult.Text = adoSet.Fields("GeomchCd").Value & ""
        ssResult.Col = 14: ssResult.Text = adoSet.Fields("DeptCode").Value & ""
        ssResult.Col = 15: ssResult.Text = adoSet.Fields("DeptNameK").Value & ""
        ssResult.Col = 16: ssResult.Text = adoSet.Fields("GeomsaCM").Value & ""
        ssResult.Col = 17: ssResult.Text = adoSet.Fields("JeobsuT1").Value & ":" & _
                                          adoSet.Fields("JeobsuT2").Value & ""
        ssResult.Col = 18: ssResult.Text = adoSet.Fields("OrderDt").Value & "'"
        
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)

End Sub

Private Sub cmdQuery6_Click()
    Dim sFrDate         As String
    Dim sToDate         As String
    Dim sWhere          As String * 1
    
    
    Call SpreadSetClear(ssResult)
    
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    If optWhere(0).Value = True Then sWhere = "R"
    If optWhere(1).Value = True Then sWhere = "P"
    If optWhere(2).Value = True Then sWhere = "C'"
    
'o  strSql = ""
'o  strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INDEX_PATIENT0) */"

    strSql = ""
    strSql = strSql & " SELECT a.*, a.RowID,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDT, 'YYYY-MM-DD') JeobsuDt, "
    strSql = strSql & "        TO_CHAR(a.OrderDt,  'YYYY-MM-DD') OrderDt,  "
    strSql = strSql & "        TO_CHAR(a.GeomsaDt, 'YYYY-MM-DD') GeomsaDt, "
    strSql = strSql & "        b.Sname,       b.AgeYY, b.Sex,  c.Deptnamek, d.Drname, e.Name, f.Codenm"
    strSql = strSql & " FROM   TWEXAM_GENERAL  a, "
    strSql = strSql & "        TWEXAM_IDNOMST  b, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT      c, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR    d, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PASS      e, "
    strSql = strSql & "        TWEXAM_Sample   f  "
    strSql = strSql & " WHERE  a.JeobsuDt >= TO_DATE('" & sFrDate & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.JeobsuDt <= TO_DATE('" & sToDate & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.Slipno1   = " & Val(Left(cmbSLip.Text, 2))
    strSql = strSql & " AND    a.Geomsaja  = '" & Left(cmbUser.Text, 6) & "'"
    strSql = strSql & " AND    a.GbCh      = 'Y'"
    strSql = strSql & " AND    a.DeptCode  != 'GE'"
    strSql = strSql & " AND    a.Status    = '" & sWhere & "'"
    strSql = strSql & " AND    a.Ptno      = b.Ptno(+)"
    strSql = strSql & " AND    a.DeptCode  = c.DeptCode(+)"
    strSql = strSql & " AND    a.DrCode    = d.DrCode(+)"
    strSql = strSql & " AND    a.GeomchCD  = f.Code(+)"
    strSql = strSql & " AND    a.Geomsaja  = e.Idnumber(+)"
    strSql = strSql & " AND   (e.Programid = ' ' OR e.Programid IS NULL)"
    strSql = strSql & " ORDER  BY a.Jeobsudt, a.SLipno2"
    
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        ssResult.Row = ssResult.DataRowCnt + 1
        ssResult.Col = 1:
        ssResult.Col = 2:  ssResult.Text = adoSet.Fields("JeobsuDt").Value & ""
        ssResult.Col = 3:  ssResult.Text = adoSet.Fields("SLipno1").Value & ""
        ssResult.Col = 4:  ssResult.Text = adoSet.Fields("SLipno2").Value & ""
        ssResult.Col = 5:  ssResult.Text = adoSet.Fields("RoomCode").Value & ""
        ssResult.Col = 6:  ssResult.Text = adoSet.Fields("Ptno").Value & ""
        ssResult.Col = 7:  ssResult.Text = adoSet.Fields("Sname").Value & ""
        ssResult.Col = 8:  ssResult.Text = adoSet.Fields("Sex").Value & ""
        ssResult.Col = 9:  ssResult.Text = adoSet.Fields("AgeYY").Value & ""
        ssResult.Col = 10: ssResult.Text = adoSet.Fields("GeomsaDt").Value & ""
        ssResult.Col = 11: ssResult.Text = adoSet.Fields("GeomsaT1").Value & ":" & _
                                          adoSet.Fields("GeomsaT2").Value & ""
        ssResult.Col = 12: ssResult.Text = adoSet.Fields("Geomsaja").Value & ""
        ssResult.Col = 13: ssResult.Text = adoSet.Fields("GeomchCd").Value & ""
        ssResult.Col = 14: ssResult.Text = adoSet.Fields("DeptCode").Value & ""
        ssResult.Col = 15: ssResult.Text = adoSet.Fields("DeptNameK").Value & ""
        ssResult.Col = 16: ssResult.Text = adoSet.Fields("GeomsaCM").Value & ""
        ssResult.Col = 17: ssResult.Text = adoSet.Fields("JeobsuT1").Value & ":" & _
                                          adoSet.Fields("JeobsuT2").Value & ""
        ssResult.Col = 18: ssResult.Text = adoSet.Fields("OrderDt").Value & "'"
        
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)


End Sub

Private Sub cmdQuery7_Click()
    Dim sFrDate         As String
    Dim sToDate         As String
    Dim sWhere          As String * 1
    
    
    Call SpreadSetClear(ssResult)
    
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    If optWhere(0).Value = True Then sWhere = "R"
    If optWhere(1).Value = True Then sWhere = "P"
    If optWhere(2).Value = True Then sWhere = "C'"
    
'o  strSql = ""
'o  strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INDEX_PATIENT0) */"

    strSql = ""
    strSql = strSql & " SELECT a.*, a.RowID,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDT, 'YYYY-MM-DD') JeobsuDt, "
    strSql = strSql & "        TO_CHAR(a.OrderDt,  'YYYY-MM-DD') OrderDt,  "
    strSql = strSql & "        TO_CHAR(a.GeomsaDt, 'YYYY-MM-DD') GeomsaDt, "
    strSql = strSql & "        b.Sname,       b.AgeYY, b.Sex,  c.Deptnamek, d.Drname, e.Name, f.Codenm"
    strSql = strSql & " FROM   TWEXAM_GENERAL  a, "
    strSql = strSql & "        TWEXAM_IDNOMST  b, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT      c, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR    d, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PASS      e, "
    strSql = strSql & "        TWEXAM_Sample   f  "
    strSql = strSql & " WHERE  a.JeobsuDt >= TO_DATE('" & sFrDate & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.JeobsuDt <= TO_DATE('" & sToDate & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.Slipno1   = " & Val(Left(cmbSLip.Text, 2))
    strSql = strSql & " AND    a.GbCh      = 'Y'"
    strSql = strSql & " AND    a.DeptCode  != 'GE'"
    strSql = strSql & " AND    a.Status    = '" & sWhere & "'"
    strSql = strSql & " AND    a.Ptno      = b.Ptno(+)"
    strSql = strSql & " AND    a.DeptCode  = c.DeptCode(+)"
    strSql = strSql & " AND    a.DrCode    = d.DrCode(+)"
    strSql = strSql & " AND    a.GeomchCD  = '" & Left(cmbSample.Text, 8) & "'"
    strSql = strSql & " AND    a.GeomchCD  = f.Code(+)"
    strSql = strSql & " AND    a.Geomsaja  = e.Idnumber(+)"
    strSql = strSql & " AND   (e.Programid = ' ' OR e.Programid IS NULL)"
    strSql = strSql & " ORDER  BY a.JeobsuDt, a.SLipno2"
    
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        ssResult.Row = ssResult.DataRowCnt + 1
        ssResult.Col = 1:
        ssResult.Col = 2:  ssResult.Text = adoSet.Fields("JeobsuDt").Value & ""
        ssResult.Col = 3:  ssResult.Text = adoSet.Fields("SLipno1").Value & ""
        ssResult.Col = 4:  ssResult.Text = adoSet.Fields("SLipno2").Value & ""
        ssResult.Col = 5:  ssResult.Text = adoSet.Fields("RoomCode").Value & ""
        ssResult.Col = 6:  ssResult.Text = adoSet.Fields("Ptno").Value & ""
        ssResult.Col = 7:  ssResult.Text = adoSet.Fields("Sname").Value & ""
        ssResult.Col = 8:  ssResult.Text = adoSet.Fields("Sex").Value & ""
        ssResult.Col = 9:  ssResult.Text = adoSet.Fields("AgeYY").Value & ""
        ssResult.Col = 10: ssResult.Text = adoSet.Fields("GeomsaDt").Value & ""
        ssResult.Col = 11: ssResult.Text = adoSet.Fields("GeomsaT1").Value & ":" & _
                                          adoSet.Fields("GeomsaT2").Value & ""
        ssResult.Col = 12: ssResult.Text = adoSet.Fields("Geomsaja").Value & ""
        ssResult.Col = 13: ssResult.Text = adoSet.Fields("GeomchCd").Value & ""
        ssResult.Col = 14: ssResult.Text = adoSet.Fields("DeptCode").Value & ""
        ssResult.Col = 15: ssResult.Text = adoSet.Fields("DeptNameK").Value & ""
        ssResult.Col = 16: ssResult.Text = adoSet.Fields("GeomsaCM").Value & ""
        ssResult.Col = 17: ssResult.Text = adoSet.Fields("JeobsuT1").Value & ":" & _
                                          adoSet.Fields("JeobsuT2").Value & ""
        ssResult.Col = 18: ssResult.Text = adoSet.Fields("OrderDt").Value & "'"
        
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)

End Sub

Private Sub cmdSelect_Click()
    
    If cmdSelect.Caption = "전체선택" Then
        For i = 1 To ssResult.DataRowCnt
            ssResult.Row = i
            ssResult.Col = 1
            ssResult.Value = True
        Next
        cmdSelect.Caption = "전체해제"
        
    Else
        For i = 1 To ssResult.DataRowCnt
            ssResult.Row = i
            ssResult.Col = 1
            ssResult.Value = False
        Next
        cmdSelect.Caption = "전체선택"
    End If
    
End Sub

Private Sub Form_Load()
    Dim sDeptC          As String * 4
    Dim sWardC          As String * 4
    Dim sFrDate         As String
    Dim sToDate         As String
    
    
    GoSub Clear_Forma_TextBox
    
    dtFrDate.Value = Dual_Date_Get("yyyy-MM-dd")
    dtToDate.Value = Dual_Date_Get("yyyy-MM-dd")
    
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    txtFrJeobsuT1.Text = "00"
    txtFrJeobsuT2.Text = "00"
    
    txtToJeobsuT1.Text = Dual_Date_Get("hh24")
    txtToJeobsuT2.Text = Dual_Date_Get("mi")
    
    
    GoSub SLip_Select
    
    GiExamNumb = Val(GetSetting("CP", "CPRESULT", "SLip"))
    Call SetComboBox(cmbSLip, GiExamNumb, 2)
    
    GoSub SELECT_Query_Dept
    GoSub SELECT_Query_Ward
    GoSub SELECT_Query_User
    GoSub SELECT_Query_Sample
    
    Exit Sub
    
Clear_Forma_TextBox:
    For i = 0 To Me.Count - 1
        If TypeOf Me.Controls(i) Is VB.TextBox Then
            Me.Controls(i).Text = ""
        End If
    Next
    Return
    
SELECT_Query_Dept:
'o  strSql = ""
'o  strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_DEPT INDEX_DEPT0) */"

    strSql = ""
    strSql = strSql & " SELECT a.DeptCode, b.DeptNamek"
    strSql = strSql & " FROM   TWEXAM_General a,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT     b"
    strSql = strSql & " WHERE  a.JeobsuDt >= TO_DATE('" & sFrDate & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.JeobsuDt <= TO_DATE('" & sToDate & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.Deptcode  = b.Deptcode(+)"
    strSql = strSql & " GROUP  BY a.DeptCode, b.DeptNamek"
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    cmbDept.Clear
    
    Do Until adoSet.EOF
        sDeptC = adoSet.Fields("DeptCode").Value & ""
        cmbDept.AddItem sDeptC & "." & Trim(adoSet.Fields("DeptNamek").Value & "")
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return
    
SELECT_Query_Ward:
'o  strSql = ""
'o  strSql = strSql & "  SELECT /*+ INDEX (TWBas_Room INDEX_Room0) */"

    strSql = ""
    strSql = strSql & " SELECT  c.WardCode, c.WardName"
    strSql = strSql & "  FROM   TWEXAM_General a,        "
    strSql = strSql & "         TW_MIS_PMPA.TWBas_Room     b,        "
    strSql = strSql & "         TW_MIS_PMPA.TWBas_Ward     c  "
    strSql = strSql & "  WHERE  a.JeobsuDt >= TO_DATE('" & sFrDate & "','yyyy-MM-dd')"
    strSql = strSql & "  AND    a.JeobsuDt <= TO_DATE('" & sToDate & "','yyyy-MM-dd')"
    strSql = strSql & "  AND    a.RoomCode = b.Roomcode(+) "
    strSql = strSql & "  AND    b.WardCode = c.WardCode(+) "
    strSql = strSql & "  GROUP  BY c.WardCode, c.WardName"
    
    cmbWard.Clear
    If False = adoSetOpen(strSql, adoSet) Then Return
    Do Until adoSet.EOF
        sWardC = adoSet.Fields("WardCode").Value & ""
        cmbWard.AddItem sWardC & "." & Trim(adoSet.Fields("WardName").Value & "")
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
        
    Return
    
    
    
SELECT_Query_User:
    Dim sUserID     As String * 6
    
    strSql = ""
    strSql = strSql & "  SELECT b.IDNumber, b.Name"
    strSql = strSql & "  FROM   TWEXAM_General a,"
    strSql = strSql & "         TW_MIS_PMPA.TWBAS_PASS     b "
    strSql = strSql & "  WHERE  a.JeobsuDt >= TO_DATE('" & sFrDate & "','yyyy-MM-dd')"
    strSql = strSql & "  AND    a.JeobsuDt <= TO_DATE('" & sToDate & "','yyyy-MM-dd')"
    strSql = strSql & "  AND    a.Geomsaja  = b.IdNumber(+)"
    strSql = strSql & "  AND   (b.PRogramid = ' ' OR b.PRogramid is null)"
    strSql = strSql & "  GROUP  BY b.IDNumber, b.Name"
    
    cmbUser.Clear
    If False = adoSetOpen(strSql, adoSet) Then Return
    Do Until adoSet.EOF
        sUserID = adoSet.Fields("Idnumber").Value & ""
        cmbUser.AddItem sUserID & "." & Trim(adoSet.Fields("name").Value & "")
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return
    

SELECT_Query_Sample:
    Dim sSampleCode     As String * 8
    
    strSql = ""
    strSql = strSql & "  SELECT a.GeomchCD, b.Codenm"
    strSql = strSql & "  FROM   TWEXAM_General a,"
    strSql = strSql & "         TWEXAM_Sample  b "
    strSql = strSql & "  WHERE  a.JeobsuDt >= TO_DATE('" & sFrDate & "','yyyy-MM-dd')"
    strSql = strSql & "  AND    a.JeobsuDt <= TO_DATE('" & sToDate & "','yyyy-MM-dd')"
    strSql = strSql & "  AND    a.GeomchCD  = b.Code(+)"
    strSql = strSql & "  GROUP  BY a.GeomchCD, b.Codenm"
    
    cmbSample.Clear
    If False = adoSetOpen(strSql, adoSet) Then Return
    Do Until adoSet.EOF
        sSampleCode = adoSet.Fields("GeomchCD").Value & ""
        cmbSample.AddItem sSampleCode & "." & Trim(adoSet.Fields("CodeNM").Value & "")
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return
    
    Return
    
    
SLip_Select:
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_Specode"
    strSql = strSql & " WHERE  Codegu = '12'"
    strSql = strSql & " AND    Codeky < '90'"
'    strSql = strSql & " AND    Codeky < '52'"
    strSql = strSql & " ORDER  BY Codeky"
    
    cmbSLip.Clear
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        cmbSLip.AddItem Trim(adoSet.Fields("Codeky").Value & "") & ". " & _
                        Trim(adoSet.Fields("Codenm").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
        
    Return

End Sub

Private Sub ssResult_DblClick(ByVal Col As Long, ByVal Row As Long)

    If Row = 0 Then
        ssResult.Col = 1
        ssResult.Col2 = ssResult.MaxCols
        ssResult.Row = 1
        ssResult.Row2 = ssResult.DataRowCnt
        
        ssResult.SortBy = SS_SORT_BY_ROW
        ssResult.SortKey(1) = Col
        
        If ssResult.SortKeyOrder(1) = SortKeyOrderAscending Then
            ssResult.SortKeyOrder(1) = SortKeyOrderDescending
        Else
            ssResult.SortKeyOrder(1) = SortKeyOrderAscending
        End If
        ssResult.Action = SS_ACTION_SORT
    End If
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1: Unload Me
        Case 3: GoSub Form_Clear_Sub
    End Select
    Exit Sub
    

Form_Clear_Sub:
    Screen.MousePointer = vbHourglass
    
    Call Form_Load
    Call SpreadSetClear(ssResult)
    vaTabPro1.ActiveTab = 0
    optWhere(2).SetFocus
    
    For i = 1 To ssResult.DataRowCnt
        ssResult.Row = i
        ssResult.Col = 1
        ssResult.Value = False
    Next
    cmdSelect.Caption = "전체선택"
    
    Screen.MousePointer = vbDefault
    
    Return
    
    
End Sub

Private Sub txtFrJeobsuT1_GotFocus()
    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.Text)
End Sub

Private Sub txtFrJeobsuT1_LostFocus()
    
    txtFrJeobsuT1.Text = Format(txtFrJeobsuT1.Text, "00")
    
End Sub

Private Sub txtFrJeobsuT2_GotFocus()
    
    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.Text)

End Sub

Private Sub txtFrJeobsuT2_LostFocus()
    
    txtFrJeobsuT2.Text = Format(txtFrJeobsuT2.Text, "00")
    
End Sub

Private Sub txtLabno_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txtLabno.Text = Format(txtLabno.Text, "00000")
        txtToLabno.SetFocus
        
    End If
    
End Sub

Private Sub txtPtno_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txtPtno.Text = Format(txtPtno.Text, "00000000")
        cmdQuery1.SetFocus
    End If
    
End Sub

Private Sub txtSname_GotFocus()
    
    txtSname.SelStart = 0
    txtSname.SelLength = Len(txtSname.Text)
    
    txtSname.IMEMode = vbIMEModeHangul
    
End Sub

Private Sub txtSname_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        cmdQuery2.SetFocus
    End If
    
End Sub

Private Sub txtToJeobsuT1_GotFocus()

    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.Text)

End Sub

Private Sub txtToJeobsuT1_LostFocus()
    
    txtToJeobsuT1.Text = Format(txtToJeobsuT1.Text, "00")
    
End Sub

Private Sub txtToJeobsuT2_GotFocus()

    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.Text)

End Sub

Private Sub txtToJeobsuT2_LostFocus()

    txtToJeobsuT2.Text = Format(txtToJeobsuT2.Text, "00")
    
End Sub

Private Sub txtToLabno_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Trim(txtToLabno.Text) = "" Then
            txtToLabno.Text = txtLabno.Text
        Else
            txtToLabno.Text = Format(txtToLabno.Text, "00000")
        End If
        cmdQuery3.SetFocus
    End If
    
End Sub

Private Sub txtToLabno_LostFocus()

    If Trim(txtToLabno.Text) = "" Then
        txtToLabno.Text = txtLabno.Text
    Else
        txtToLabno.Text = Format(txtToLabno.Text, "00000")
    End If



End Sub

