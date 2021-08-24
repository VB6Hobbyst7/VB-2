VERSION 5.00
Begin VB.Form frmSetup 
   Caption         =   "File Path Configuration"
   ClientHeight    =   735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   735
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   255
      Left            =   5070
      TabIndex        =   3
      Top             =   30
      Width           =   1185
   End
   Begin VB.CommandButton cmdPath 
      Caption         =   "Change Dir"
      Height          =   255
      Left            =   3900
      TabIndex        =   2
      Top             =   30
      Width           =   1155
   End
   Begin VB.TextBox txtPath 
      Height          =   345
      Left            =   60
      TabIndex        =   1
      Top             =   300
      Width           =   6165
   End
   Begin VB.Label Label1 
      Caption         =   "[Result XML File Path]"
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2505
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPath_Click()
   
    Dim lsPath As String
    lsPath = BrowsePath
    If Trim(lsPath) <> "" Then
        txtPath = lsPath
    End If
    
End Sub

Private Sub cmdSave_Click()
    gFilePath = txtPath.Text
    
    Call WritePrivateProfileString("config", "FilePath", gFilePath, App.Path & "\Interface.ini")
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    txtPath.Text = gFilePath
    
End Sub


Public Function BrowsePath(Optional ByVal bDontShowBrowser As Boolean) As String
   Dim lRet    As Long
   Dim sPath   As String

   sPath = GetInitEntry("Main", "Last Path", txtPath)
   If Not bDontShowBrowser Then
        sPath = BrowseForFolder(Me.hWnd, "Select a Folder with Images...", sPath)
   End If
   If Len(sPath) > 0 Then
      BrowsePath = sPath
      Else
      BrowsePath = ""
   End If
End Function

