VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCalendar 
   Caption         =   "´Þ·Â"
   ClientHeight    =   2820
   ClientLeft      =   4680
   ClientTop       =   2505
   ClientWidth     =   2730
   BeginProperty Font 
      Name            =   "±¼¸²Ã¼"
      Size            =   12
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2820
   ScaleWidth      =   2730
   Begin MSComCtl2.MonthView MonthDate 
      Height          =   2820
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   4974
      _Version        =   393216
      ForeColor       =   8421376
      BackColor       =   8421376
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   1
      MonthBackColor  =   13828095
      StartOfWeek     =   24510465
      TitleBackColor  =   8421440
      TitleForeColor  =   16777215
      CurrentDate     =   36145
   End
End
Attribute VB_Name = "FrmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    MonthDate.Value = Now
    
End Sub

Private Sub MonthDate_DateDblClick(ByVal DateDblClicked As Date)

    GstrDate = Format(MonthDate.Value, "YYYY-MM-DD")
    
    Unload Me
    
End Sub
