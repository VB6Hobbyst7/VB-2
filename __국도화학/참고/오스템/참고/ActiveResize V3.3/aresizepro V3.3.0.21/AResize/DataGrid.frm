VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form DataGrid 
   Caption         =   "Data Grid Resize Demo"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8520
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   5820
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Do Not Resize Grid Columns And Font... Set This Before You Resize The Form!"
      Height          =   195
      Left            =   90
      TabIndex        =   7
      Top             =   5040
      Width           =   6555
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   180
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   18
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      AutoCenterForm  =   -1  'True
      FormHeightDT    =   6225
      FormWidthDT     =   8640
      FormScaleHeightDT=   5820
      FormScaleWidthDT=   8520
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4155
      Left            =   0
      TabIndex        =   4
      Top             =   810
      Width           =   8505
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3930
         Left            =   60
         TabIndex        =   5
         Top             =   150
         Width           =   8370
         _ExtentX        =   14764
         _ExtentY        =   6932
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   435
      Index           =   3
      Left            =   6630
      TabIndex        =   0
      Top             =   5340
      Width           =   1845
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete Customer"
      Height          =   435
      Index           =   2
      Left            =   3840
      TabIndex        =   3
      Top             =   5340
      Width           =   1845
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Edit Customer"
      Height          =   435
      Index           =   1
      Left            =   1950
      TabIndex        =   2
      Top             =   5340
      Width           =   1845
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Customer"
      Height          =   435
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   5340
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Data Grid Resize Demo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   270
      Width           =   4395
   End
End
Attribute VB_Name = "DataGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'-------------------------------------------------------------------------
'                   Not even a single line of code!!!
'-------------------------------------------------------------------------
    

'The following code is used to populate the grid with data from database
    
Dim cmd As ADODB.Command
Dim rs As ADODB.Recordset

Private Sub Check1_Click()
    'Tell ActiveResize not to resize the grid columns and fonts, so that
    'more columns become visible when the grid becomes wider.
    'This will take effect immediately from the current state of the grid,
    'i.e if the grid columns are currently wide because you have maximized
    'the form and you tick this checkbox and make the form smaller, the
    'columns and fonts will not get smaller.
    'You can specify this at design-time by inserting Ex_Columns,Ex_Font
    'in the Tag property of the grid
    If Check1.Value = 1 Then
        DataGrid1.Tag = "Ex_Columns,Ex_Font"
    Else
        DataGrid1.Tag = ""
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    If Index = 3 Then Unload Me
End Sub

Private Sub Form_Load()

    'Load data from database to the grid
    DBName = "\Sample.mdb"
    ConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & DBName
    
    Set cmd = New ADODB.Command
    Set rs = New ADODB.Recordset
    
    With cmd
        .ActiveConnection = ConStr
        .ActiveConnection.CursorLocation = adUseClient
        .CommandType = adCmdText
        .CommandText = "SELECT * FROM Customers"
        Set rs = .Execute
    End With
        
    Set DataGrid1.DataSource = rs
    DataGrid1.Columns(0).Width = 1500
    DataGrid1.Columns(1).Width = 2500
    DataGrid1.Columns(3).Width = 2500
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Clean up
    Set DataGrid1.DataSource = Nothing
    rs.Close
    Set rs = Nothing
    cmd.ActiveConnection.Close
    Set cmd = Nothing

End Sub
