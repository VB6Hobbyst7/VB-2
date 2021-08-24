VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl urcStkList 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin MSComctlLib.ImageList imgList 
      Left            =   3810
      Top             =   690
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "urcStkList.ctx":0000
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "urcStkList.ctx":08DA
            Key             =   "open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "urcStkList.ctx":11B4
            Key             =   "choice"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView trvList 
      Height          =   2535
      Left            =   420
      TabIndex        =   0
      Top             =   360
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   4471
      _Version        =   393217
      Style           =   7
      ImageList       =   "imgList"
      Appearance      =   0
   End
End
Attribute VB_Name = "urcStkList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event Click()
Event DblClick()
Event Expand(Node As MSComctlLib.Node)
Event Collapse(Node As MSComctlLib.Node)

Private Sub trvList_Click()

    RaiseEvent Click

End Sub

Private Sub trvList_Collapse(ByVal Node As MSComctlLib.Node)

    RaiseEvent Collapse(Node)

End Sub

Private Sub trvList_DblClick()

    RaiseEvent DblClick

End Sub

Private Sub trvList_Expand(ByVal Node As MSComctlLib.Node)

    RaiseEvent Expand(Node)

End Sub

Private Sub UserControl_Initialize()

    trvList.Nodes.Clear
    gSql = "select * from mstSTKG where reagentfg = 1 order by kindnm "
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            While (Not .EOF)
                trvList.Nodes.Add , , .Fields("kindcd").Value, .Fields("kindnm").Value, "close", "open"
                
                .MoveNext
            Wend
            .Close
        End If
    End With

End Sub

Private Sub UserControl_Resize()

    trvList.Top = 0
    trvList.Left = 0
    trvList.Width = UserControl.Width
    trvList.Height = UserControl.Height

End Sub
