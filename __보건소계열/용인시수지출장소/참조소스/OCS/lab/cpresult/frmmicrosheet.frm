VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMicroSheet 
   Caption         =   "Form1"
   ClientHeight    =   6840
   ClientLeft      =   540
   ClientTop       =   1680
   ClientWidth     =   11115
   BeginProperty Font 
      Name            =   "±º∏≤√º"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   11115
   WindowState     =   2  '√÷¥Î»≠
   Begin FPSpreadADO.fpSpread sprSelect 
      Height          =   5730
      Left            =   270
      TabIndex        =   3
      Top             =   765
      Width           =   4380
      _Version        =   196608
      _ExtentX        =   7726
      _ExtentY        =   10107
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±º∏≤√º"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   10
      ScrollBars      =   2
      SpreadDesigner  =   "frmMicroSheet.frx":0000
      UserResize      =   0
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker dtFrDate 
      Height          =   330
      Left            =   1215
      TabIndex        =   1
      Top             =   315
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   24510467
      CurrentDate     =   36566
   End
   Begin MSComCtl2.DTPicker dtToDate 
      Height          =   330
      Left            =   2700
      TabIndex        =   2
      Top             =   315
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   24510467
      CurrentDate     =   36566
   End
   Begin MSForms.CommandButton cmdQuery 
      Height          =   420
      Left            =   4365
      TabIndex        =   4
      Top             =   315
      Width           =   1590
      Size            =   "2805;741"
      FontName        =   "±º∏≤√º"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      Caption         =   "¡¢ºˆ¿œ¿⁄:"
      BeginProperty Font 
         Name            =   "±º∏≤√º"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   225
      TabIndex        =   0
      Top             =   360
      Width           =   915
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmMicroSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQuery_Click()
    
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_ItemML"
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub
