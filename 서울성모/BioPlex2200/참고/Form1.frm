VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton Command2 
      Caption         =   "ȣ��"
      Height          =   615
      Left            =   5460
      TabIndex        =   2
      Top             =   1890
      Width           =   1725
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȣ��"
      Height          =   615
      Left            =   5400
      TabIndex        =   1
      Top             =   1080
      Width           =   1725
   End
   Begin VB.TextBox Text1 
      Height          =   435
      Left            =   270
      TabIndex        =   0
      Text            =   "http://his012edu.cmcnu.or.kr/himed/webapps/com/commonweb/xrw/.live?"
      Top             =   210
      Width           =   6945
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim sUrl, sPost, sParam As String
Dim sRcvData, sData As String

    Dim sHtmlLine
    Dim hInet As Object
    

    Set hInet = CreateObject("InetCtls.Inet")  'Inet ��ü ����

 

    'sUrl = "http://********"
    'http://his012edu.cmcnu.or.kr/himed/webapps/com/commonweb/xrw/.live?
    sUrl = Text1.Text
    'ex1) �α��λ�� :
    '     >> http://his012edu.cmcnu.or.kr/himed/webapps/com/commonweb/xrw/.live?submit_id=TRLII00104&business_id=lis&ex_interface=12345678|01
    'sPost = "Param1=" & Replace(sData, "&", "%26") '������ &�� ó��
            sPost = "submit_id=" & "TRLII00104" & "&"  '������ &�� ó��
    sPost = sPost & "business_id=" & "lis" & "&" '������ &�� ó��
    sPost = sPost & "ex_interface=" & "12345678|01" & "&" '������ &�� ó��
    
'    ' POST ������� �������� ȣ��
'    hInet.Execute sUrl, "POST", sPost, "Content-Type:application/x-www-form-urlencoded"
    
    ' GET ������� �������� ȣ��
   ' hInet.execute "http://his012edu.cmcnu.or.kr/himed/webapps/com/commonweb/xrw/.live?submit_id=TRLII00103&business_id=lis&ex_interface=12345678|01&" ', "Content-Type:application/x-www-form-urlencoded"
    
'    hInet.execute "http://his012edu.cmcnu.or.kr/himed/webapps/com/commonweb/xrw/.live?submit_id=TRLII00104&business_id=lis&ex_interface=�α��λ��"
    
    
    
'    hInet.execute "http://his012edu.cmcnu.or.kr/himed/webapps/com/commonweb/xrw/.live?submit_id=TRLII00104&business_id=lis&ex_interface=93031007|01"
    
    hInet.execute "http://his012edu.cmcnu.or.kr/himed/webapps/com/commonweb/xrw/.live?submit_id=TRLII00104&business_id=lis&ex_interface=93031007|01&instcd=012&userid=93031007", "Content-Type:application/x-www-form-urlencoded"
    
    While hInet.StillExecuting
        DoEvents
    Wend
    sData = hInet.GetChunk(1024, sHtmlLine)


     Do While LenB(sData) > 0
        sRcvData = sRcvData & sData
        sData = hInet.GetChunk(1024, sHtmlLine)
    Loop


End Sub
