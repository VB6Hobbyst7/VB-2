Attribute VB_Name = "modUIMasterStatics"
Option Explicit


Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long


'Private fB800 As frmBBS800
'Private fB801 As frmBBS801
'Private fB802 As frmBBS802
'Private fB803 As frmBBS803
'
'Private fB811 As frmBBS811
'Private fB812 As frmBBS812
'Private fB813 As frmBBS813
'Private fB814 As frmBBS814
'Private fB815 As frmBBS815
'Private fB816 As frmBBS816
'Private fB817 As frmBBS817
'Private fB818 As frmBBS818
'Private fB819 As frmBBS819
'Private fB820 As frmBBS820
'Private fB821 As frmBBS821
'Private fB822 As frmBBS822
'Private fB823 As frmBBS823
'
'Private fB861B As frmBBS861
'Private fB861C As frmBBS861
'Private fB861D As frmBBS861
'Private fB861E As frmBBS861
'Private fB861F As frmBBS861
'Private fB861G As frmBBS861
'Private fB861H As frmBBS861
'Private fB861I As frmBBS861
'Private fB861J As frmBBS861
'Private fB861K As frmBBS861
'Private fB861L As frmBBS861
'Private fB861M As frmBBS861
'Private fB861N As frmBBS861
'
'
'
'





'---------------------------------------------------------------------------------------------
'  ���,��� UI
'---------------------------------------------------------------------------------------------
Public Sub StaticsClose()
   Dim tmpForm As Form

   For Each tmpForm In Forms
      With tmpForm
        If UCase(Mid(.name, 1, 7)) = "FRMBBS9" Then
            Unload tmpForm
            Set tmpForm = Nothing
        End If
      End With
   Next
   
   
   
   For Each tmpForm In Forms
      With tmpForm
        If UCase(.name) = "FRMSTATICS" Then
            Unload tmpForm
            Set tmpForm = Nothing
        End If
      End With
   Next
End Sub

Public Sub StaticsTreeviewLoad(tvwMenu As Object)
    With tvwMenu
        .Nodes.Clear
        
        Call .Nodes.Add(, , "B1", "����", 1)
        Call .Nodes.Add("B1", tvwChild, "BBS925", "�����ϸ���", 2)
        Call .Nodes.Add("B1", tvwChild, "BBS924", "���������������Ȳ", 2)
        .Nodes(.Nodes.Count).EnsureVisible
        
        Call .Nodes.Add("B1", tvwNext, "B2", "���׿� ����Ʈ", 1)
        Call .Nodes.Add("B2", tvwChild, "BBS920", "�����������̼�", 2)
        Call .Nodes.Add("B2", tvwChild, "BBS921", "��������������", 2)
        Call .Nodes.Add("B2", tvwChild, "BBS922", "���������ȸ��", 2)
        Call .Nodes.Add("B2", tvwChild, "BBS923", "�����ڴ���", 2)
        .Nodes(.Nodes.Count).EnsureVisible
        
        Call .Nodes.Add("B2", tvwNext, "B3", "���", 1)
        Call .Nodes.Add("B3", tvwChild, "BBS911", "MSBOS �ۼ�", 2)
        Call .Nodes.Add("B3", tvwChild, "BBS912", "ȯ�ں� ��������", 2)
        Call .Nodes.Add("B3", tvwChild, "BBS913", "�����Ϻ�", 2)
        Call .Nodes.Add("B3", tvwChild, "BBS914", "C-T Ratio", 2)
        Call .Nodes.Add("B3", tvwChild, "BBS915", "�������� �Ǽ�", 2)
        Call .Nodes.Add("B3", tvwChild, "BBS916", "�������ۿ� �Ǽ�", 2)
        Call .Nodes.Add("B3", tvwChild, "BBS917", "�������ۿ� ȯ�ڸ���Ʈ", 2)
        .Nodes(.Nodes.Count).EnsureVisible
        
        Call .Nodes.Add("B3", tvwNext, "B4", "��ȸ/���", 1)
        Call .Nodes.Add("B4", tvwChild, "BBS961", "���� ��ȸ", 2)
        .Nodes(.Nodes.Count).EnsureVisible
        
        .BorderStyle = vbFixedSingle
    End With
End Sub
'
Public Sub StaticsTreeviewCollapse(ByVal Node As MSComctlLib.Node)
   Node.Image = "Close"
End Sub

Public Sub StaticsTreeviewExpand(ByVal Node As MSComctlLib.Node)
   Node.Image = "Open"
End Sub

Public Sub StaticsTreeviewNodeClick(ByVal Node As MSComctlLib.Node)

On Error GoTo StaticsTreeviewNodeClick_error

    Select Case Node.Key
        Case "BBS911":    frmBBS911.Show: frmBBS911.ZOrder
        Case "BBS912":    frmBBS912.Show: frmBBS912.ZOrder
        Case "BBS913":    frmBBS913.Show: frmBBS913.ZOrder
        Case "BBS914":    frmBBS914.Show: frmBBS914.ZOrder
        Case "BBS915":    frmBBS915.Show: frmBBS915.ZOrder
        Case "BBS916":    frmBBS916.Show: frmBBS916.ZOrder
        Case "BBS917":    frmBBS917.Show: frmBBS917.ZOrder
        
        Case "BBS920":    frmBBS920.Show: frmBBS920.ZOrder
        Case "BBS921":    frmBBS921.Show: frmBBS921.ZOrder
        Case "BBS922":    frmBBS922.Show: frmBBS922.ZOrder
        Case "BBS923":    frmBBS923.Show: frmBBS923.ZOrder
        Case "BBS924":    frmBBS924.Show: frmBBS924.ZOrder
        Case "BBS925":    frmBBS925.Show: frmBBS925.ZOrder
        
        Case "BBS961":    frmBBS961.Show: frmBBS961.ZOrder
        
    End Select
    
    Exit Sub
    
StaticsTreeviewNodeClick_error:
    MsgBox Err.Description
End Sub

