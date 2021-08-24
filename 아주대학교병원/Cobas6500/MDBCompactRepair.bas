Attribute VB_Name = "MDBCompactRepair"
'*-----------------------------------------------------------------
'   2014.08.13 정효준
'   MDB 파일을 압축 및 복구 해주는 기능 (현재는 압축기능만)
'   동작전에 MDB 파일의 접속은 종료하여야 한다.
'*-----------------------------------------------------------------

Public Sub FN_Compact()
    Dim DAOdb As New DAO.DBEngine ' DAO 연결
    'interface mdb 접속 끊기
    cn.Close
    
    'Kill App.Path & "\Comp_interface.mdb"
    
    'FileCopy App.Path & "\interface.mdb", App.Path & "\Comp_interface.mdb"
    'MDB 압축을 시키며 하나 생성
    DAOdb.CompactDatabase App.Path & "\interface.mdb", App.Path & "\Comp_interface.mdb"      ' 데이터 베이스 압축 및 복구
    
    If Dir(App.Path & "\interface_bak.mdb") = "interface_bak.mdb" Then
        Kill App.Path & "\interface_bak.mdb"
    End If
    
    '기존 MDB를 이름변경
    Name App.Path & "\interface.mdb" As App.Path & "\interface_bak.mdb"
    
    '압축이 완료된 MDB를 interface.mdb로 변경한다.
    Name App.Path & "\Comp_interface.mdb" As App.Path & "\interface.mdb"
    
    'interface mdb 접속
    Connect_Local
End Sub
