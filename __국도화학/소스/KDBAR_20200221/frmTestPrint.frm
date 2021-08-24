VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmTestPrint 
   BackColor       =   &H00FFFFFF&
   Caption         =   "≈◊Ω∫∆Æ √‚∑¬"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows ±‚∫ª∞™
   Begin VB.CommandButton cmdQR 
      Caption         =   "QR√‚∑¬"
      Height          =   405
      Left            =   5190
      TabIndex        =   15
      Top             =   4320
      Width           =   1425
   End
   Begin VB.CommandButton cmdTP203C 
      Caption         =   "TP203C(ACF)  √‚∑¬"
      BeginProperty Font 
         Name            =   "∏º¿∫ ∞ÌµÒ"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5100
      TabIndex        =   14
      Top             =   3330
      Width           =   1515
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "∆˜∆Æø≠±‚"
      Height          =   525
      Left            =   5370
      TabIndex        =   12
      Top             =   1350
      Width           =   1395
   End
   Begin VB.TextBox txtSetting 
      Height          =   285
      Left            =   4830
      TabIndex        =   11
      Text            =   "9600,n,8,1"
      Top             =   930
      Width           =   1935
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   4830
      TabIndex        =   10
      Text            =   "1"
      Top             =   570
      Width           =   1935
   End
   Begin VB.TextBox txtInput9 
      BeginProperty Font 
         Name            =   "∏º¿∫ ∞ÌµÒ"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   420
      TabIndex        =   9
      Top             =   3840
      Width           =   3855
   End
   Begin VB.TextBox txtInput8 
      BeginProperty Font 
         Name            =   "∏º¿∫ ∞ÌµÒ"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   420
      TabIndex        =   8
      Top             =   3420
      Width           =   3855
   End
   Begin VB.TextBox txtInput7 
      BeginProperty Font 
         Name            =   "∏º¿∫ ∞ÌµÒ"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   420
      TabIndex        =   7
      Top             =   3000
      Width           =   3855
   End
   Begin VB.TextBox txtInput6 
      BeginProperty Font 
         Name            =   "∏º¿∫ ∞ÌµÒ"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   420
      TabIndex        =   6
      Top             =   2580
      Width           =   3855
   End
   Begin VB.TextBox txtInput5 
      BeginProperty Font 
         Name            =   "∏º¿∫ ∞ÌµÒ"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   420
      TabIndex        =   5
      Top             =   2160
      Width           =   3855
   End
   Begin VB.TextBox txtInput4 
      BeginProperty Font 
         Name            =   "∏º¿∫ ∞ÌµÒ"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   420
      TabIndex        =   4
      Top             =   1740
      Width           =   3855
   End
   Begin VB.TextBox txtInput3 
      BeginProperty Font 
         Name            =   "∏º¿∫ ∞ÌµÒ"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   420
      TabIndex        =   3
      Top             =   1320
      Width           =   3855
   End
   Begin VB.TextBox txtInput2 
      BeginProperty Font 
         Name            =   "∏º¿∫ ∞ÌµÒ"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   420
      TabIndex        =   2
      Top             =   900
      Width           =   3855
   End
   Begin VB.TextBox txtInput1 
      BeginProperty Font 
         Name            =   "∏º¿∫ ∞ÌµÒ"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   420
      TabIndex        =   1
      Top             =   480
      Width           =   3855
   End
   Begin MSCommLib.MSComm comEqp 
      Left            =   4350
      Top             =   4020
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "√‚∑¬"
      BeginProperty Font 
         Name            =   "∏º¿∫ ∞ÌµÒ"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4980
      TabIndex        =   0
      Top             =   2490
      Width           =   1515
   End
   Begin VB.Label lblComStatus 
      BackStyle       =   0  '≈ı∏Ì
      Height          =   375
      Left            =   4650
      TabIndex        =   13
      Top             =   2070
      Width           =   2175
   End
End
Attribute VB_Name = "frmTestPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOpen_Click()

    comEqp.CommPort = txtPort.Text
    comEqp.RTSEnable = True
    comEqp.DTREnable = True
    comEqp.Settings = txtSetting.Text

    If comEqp.PortOpen = False Then
        comEqp.PortOpen = True
    End If

    If comEqp.PortOpen Then
        lblComStatus.Caption = "COM" & comEqp.CommPort & "∆˜∆Æ ø¨∞·º∫∞¯"
    Else
        lblComStatus.Caption = "COM" & comEqp.CommPort & "∆˜∆Æ ø¨∞·Ω«∆–"
    End If
    
End Sub

Private Sub cmdPrint_Click()
    Dim pString     As String
    
    
    'sLabel = "^XA^MD10^LH0,0^FS" + PNM+PID+ADT+BCD+AGE+OCD+DPT+SPC+GBN+PDT + "^XZ"


    'Frm_Main.Mcom.Output = "^XA"
    'Frm_Main.Mcom.Output = "^PR" & Cbo_PrinterSpeed.Text & "^FS"
    'Frm_Main.Mcom.Output = "^LH" & Txt_CenterX.Text & "," & Txt_CenterY.Text & "^FS"

    
    
'^XA
'^FO 20,80^BC N,50,Y,N,N,Y^FD 2Y0P03D9O1402P10120000^FS
'^FO 20,170^A2N,25,12^FD Name : TP203C(ACF)^FS
'^FO 20,210^A1N,25,12^FD Size : 1.5mm x 2000cm/Reel^FS
'^FO 20,250^A1N,25,12^FD Production Date : 2020.02.13^FS
'^FO 20,290^A1N,25,12^FD Expiration Date : 2020.07.13^FS
'^FO 20,330^A1N,25,12^FD Storage Temperature : -10 ~ 5°… ^FS
'^FO 20,370^A1N,25,12^FD SDI ACF Lot : JOF142Y4CM (P101) ^FS
'^FO 20,420^A1N,25,12^FD Material Code : 6906B0001D000 ^FS
'^XZ

    
    '^XA : Opening Bracket¿∏∑Œ Format¿« Ω√¿€¿ª æÀ∏∞¥Ÿ.
    '^FO : ¿Œº‚ «“ «◊∏Ò¿« ¿Œº‚ «“ ¿ßƒ°(X√‡,Y√‡)∏¶ ¡§¿««—¥Ÿ.

    
    pString = ""
    pString = pString & "^XA" & vbLf
    pString = pString & "^SEE:UHANGUL.DAT^FS" & vbLf
    pString = pString & "^PON^FS" & vbLf
    pString = pString & "^CW1,E:KFONT15.FNT^FS" & vbLf
    pString = pString & "^FO030,30^CI26^A1N,25,20^FD" & txtInput1.Text & "^FS" & vbLf '"sRcpNm"
    pString = pString & "^FO190,30^CI26^A1N,25,20^FD" & txtInput2.Text & "^FS" & vbLf     '"cGetEqpNm(mQCS101.EQP_CD)"

    pString = pString & "^FO030,60^CI26^A1N,25,20^FD" & txtInput3.Text & "^FS" & vbLf    'cGetMtrlNm(mQCS101.EQP_CD, mQCS101.MTRL_CD)"
    pString = pString & "^FO190,60^CI26^A1N,25,20^FD" & txtInput4.Text & "^FS" & vbLf    'cGetLvlNm(mQCS101.EQP_CD, mQCS101.MTRL_CD, mQCS101.LVL_CD)

    pString = pString & "^FO030,90^CI26^A1N,25,20^FD" & txtInput5.Text & "^FS" & vbLf    'mQCS101.LOT_NO
    pString = pString & "^FO190,90^CI26^A1N,25,20^FD" & txtInput6.Text & "^FS" & vbLf    'cGetUserNm(mQCS101.REQ_ID)
    pString = pString & "^FO030,120^CI26^A1N,25,20^FD" & Format(Now, "####-##-##") & " " & txtInput7.Text & "^FS" & vbLf  '"mQCS101.REQ_SEQ"
    pString = pString & "^FO190,120^CI26^A1N,25,20^FD" & txtInput8.Text & "^FS" & vbLf  '"cGetDeptNm(mQCS101.DEPT_CD)"

    pString = pString & "^BY2,2,80" & vbLf
    pString = pString & "^FO030,160^B3N,N,,Y,N^FD" & txtInput9.Text & "^FS" & vbLf      '"strQCNo"
    pString = pString & "^PQ1,1,1,Y^FS" & vbLf
    pString = pString & "^XZ" & vbLf

    comEqp.Output = pString
    
    
    
'^XA
'^SEE:UHANGUL.DAT^FS
'^PON^FS
'^CW1,E:KFONT15.FNT^FS
'^FO030,30^CI26^A1N,25,20^FD1^FS
'^FO190,30^CI26^A1N,25,20^FD2^FS
'^FO030,60^CI26^A1N,25,20^FD3^FS
'^FO190,60^CI26^A1N,25,20^FD4^FS
'^FO030,90^CI26^A1N,25,20^FD5^FS
'^FO190,90^CI26^A1N,25,20^FD67^FS
'^FO030,120^CI26^A1N,25,20^FD4-38-74 7^FS
'^FO190,120^CI26^A1N,25,20^FD8^FS
'^BY2,2,80
'^FO030,160^B3N,N,,Y,N^FD9999^FS
'^PQ1,1,1,Y^FS
'^FO50,300^BY3^BCN,100,Y,N,N^FD>;2YOP>03D9O1402P10120000^FS
'^FO50,600^BY3^BCN,100,Y,N,N^FD>;382436>6CODE128>752375152^FS
'^XZ

    
    
End Sub


Private Sub cmdQR_Click()
    Dim pString As String
    
    pString = ""
    pString = pString & "^XA" & vbLf
    pString = pString & "^FO100,100^BQN,2,10^FDNM,AAC-42^FS" & vbLf
    pString = pString & "^XZ" & vbLf

    comEqp.Output = pString
    
End Sub

Private Sub cmdTP203C_Click()
    
    Dim pString     As String

    pString = ""
    pString = pString & "^XA"
    
'^FO100,550^BC^FD12345678^FS
    'pString = pString & "^FO 120,80^BC N,50,Y,N,N,Y^FD 2Y0P03D9O1402P10120000^FS" & vbLf
    pString = pString & "^BY 2,4,50" & vbLf 'ªÁ¿Ã¡Ó ¡∂¡§
    pString = pString & "^FO 100,80^BC^FD 2Y0P03D9O1402P10120000^FS" & vbLf
    pString = pString & "^FO 120,670^A2N,25,12^FD Name : TP203C(ACF)^FS" & vbLf
    pString = pString & "^FO 120,710^A1N,25,12^FD Size : 1.5mm x 2000cm/Reel^FS" & vbLf
    pString = pString & "^FO 120,750^A1N,25,12^FD Production Date : 2020.02.13^FS" & vbLf
    pString = pString & "^FO 120,790^A1N,25,12^FD Expiration Date : 2020.07.13^FS" & vbLf
    pString = pString & "^FO 120,830^A1N,25,12^FD Storage Temperature : -10 ~ 5°… ^FS" & vbLf
    pString = pString & "^FO 120,870^A1N,25,12^FD SDI ACF Lot : JOF142Y4CM (P101) ^FS" & vbLf
    pString = pString & "^FO 120,920^A1N,25,12^FD Material Code : 6906B0001D000 ^FS" & vbLf
    pString = pString & "^XZ" & vbLf
    
    'pString = pString & "^BY5,2,270" & vbLf
    'pString = pString & "^FO100,950^BC^FD12345678^FS" & vbLf
    
    comEqp.Output = pString

Exit Sub
    
    
    pString = ""
 pString = pString & "^XA"

 pString = pString & "^FX Top section with company logo, name and address." & vbLf
 pString = pString & "^CF0,60" & vbLf
 pString = pString & "^FO50,50^GB100,100,100^FS" & vbLf
 pString = pString & "^FO75,75^FR^GB100,100,100^FS" & vbLf
 pString = pString & "^FO88,88^GB50,50,50^FS" & vbLf
 pString = pString & "^FO220,50^FDIntershipping, Inc.^FS" & vbLf
 pString = pString & "^CF0,30" & vbLf
 pString = pString & "^FO220,115^FD1000 Shipping Lane^FS" & vbLf
 pString = pString & "^FO220,155^FDShelbyville TN 38102^FS" & vbLf
 pString = pString & "^FO220,195^FDUnited States (USA)^FS" & vbLf
 pString = pString & "^FO50,250^GB700,1,3^FS" & vbLf

 pString = pString & "^FX Second section with recipient address and permit information." & vbLf
 pString = pString & "^CFA,30" & vbLf
 pString = pString & "^FO50,300^FDJohn Doe^FS" & vbLf
 pString = pString & "^FO50,340^FD100 Main Street^FS" & vbLf
 pString = pString & "^FO50,380^FDSpringfield TN 39021^FS" & vbLf
 pString = pString & "^FO50,420^FDUnited States (USA)^FS" & vbLf
 pString = pString & "^CFA,15" & vbLf
 pString = pString & "^FO600,300^GB150,150,3^FS" & vbLf
 pString = pString & "^FO638,340^FDPermit^FS" & vbLf
 pString = pString & "^FO638,390^FD123456^FS" & vbLf
 pString = pString & "^FO50,500^GB700,1,3^FS" & vbLf

 pString = pString & "^FX Third section with barcode." & vbLf
 pString = pString & "^BY5,2,270" & vbLf
 pString = pString & "^FO100,550^BC^FD12345678^FS" & vbLf

 pString = pString & "^FX Fourth section (the two boxes on the bottom)." & vbLf
 pString = pString & "^FO50,900^GB700,250,3^FS" & vbLf
 pString = pString & "^FO400,900^GB1,250,3^FS" & vbLf
 pString = pString & "^CF0,40" & vbLf
 pString = pString & "^FO100,960^FDCtr. X34B-1^FS" & vbLf
 pString = pString & "^FO100,1010^FDREF1 F00B47^FS" & vbLf
 pString = pString & "^FO100,1060^FDREF2 BL4H8^FS" & vbLf
 pString = pString & "^CF0,190" & vbLf
 pString = pString & "^FO470,955^FDCA^FS" & vbLf

     pString = pString & "^XZ" & vbLf
    comEqp.Output = pString
End Sub




































