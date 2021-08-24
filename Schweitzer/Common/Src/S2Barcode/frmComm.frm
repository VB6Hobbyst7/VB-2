VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmComm 
   Caption         =   "Form1"
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2400
   LinkTopic       =   "Form1"
   ScaleHeight     =   1980
   ScaleWidth      =   2400
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin MSCommLib.MSComm MyComm 
      Left            =   705
      Top             =   555
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "frmComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
