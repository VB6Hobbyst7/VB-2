Attribute VB_Name = "MDBCompactRepair"
'*-----------------------------------------------------------------
'   2014.08.13 ��ȿ��
'   MDB ������ ���� �� ���� ���ִ� ��� (����� �����ɸ�)
'   �������� MDB ������ ������ �����Ͽ��� �Ѵ�.
'*-----------------------------------------------------------------

Public Sub FN_Compact()
    Dim DAOdb As New DAO.DBEngine ' DAO ����
    'interface mdb ���� ����
    cn.Close
    
    'Kill App.Path & "\Comp_interface.mdb"
    
    'FileCopy App.Path & "\interface.mdb", App.Path & "\Comp_interface.mdb"
    'MDB ������ ��Ű�� �ϳ� ����
    DAOdb.CompactDatabase App.Path & "\interface.mdb", App.Path & "\Comp_interface.mdb"      ' ������ ���̽� ���� �� ����
    
    If Dir(App.Path & "\interface_bak.mdb") = "interface_bak.mdb" Then
        Kill App.Path & "\interface_bak.mdb"
    End If
    
    '���� MDB�� �̸�����
    Name App.Path & "\interface.mdb" As App.Path & "\interface_bak.mdb"
    
    '������ �Ϸ�� MDB�� interface.mdb�� �����Ѵ�.
    Name App.Path & "\Comp_interface.mdb" As App.Path & "\interface.mdb"
    
    'interface mdb ����
    Connect_Local
End Sub
