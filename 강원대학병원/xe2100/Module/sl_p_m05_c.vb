Attribute VB_Name = "sl_p_m05"
Function sl_online_result_ul_4&(oerrmsg$, ispcid$(), iexamcode$(), iresult$(), ierrflag$(), iequipcd$(), igubun$)
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_p_m05")
        If (Socket > -1) Then
                rv% = dce_push_array(Socket, "ispcid", ispcid())
                rv% = dce_push_array(Socket, "iexamcode", iexamcode())
                rv% = dce_push_array(Socket, "iresult", iresult())
                rv% = dce_push_array(Socket, "ierrflag", ierrflag())
                rv% = dce_push_array(Socket, "iequipcd", iequipcd())
                rv% = dce_push_string(Socket, "igubun", igubun)
                dce_table = dce_submit("sl_p_m05", "sl_online_result_ul_4", Socket)
        End If
        sl_online_result_ul_4 = dce_pop_long(dce_table, "dce_result")
        rv% = dce_pop_string(dce_table, "oerrmsg", oerrmsg)
        Call dce_table_destroy(dce_table)
End Function


Function sl_areano_result_ul_4&(oerrmsg$, iexamcode$(), iresult$(), ierrflag$(), iequipcd$(), iacptdate$(), islipcode$(), ihospital$(), iareano$())
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_p_m05")
        If (Socket > -1) Then
                rv% = dce_push_array(Socket, "iexamcode", iexamcode())
                rv% = dce_push_array(Socket, "iresult", iresult())
                rv% = dce_push_array(Socket, "ierrflag", ierrflag())
                rv% = dce_push_array(Socket, "iequipcd", iequipcd())
                rv% = dce_push_array(Socket, "iacptdate", iacptdate())
                rv% = dce_push_array(Socket, "islipcode", islipcode())
                rv% = dce_push_array(Socket, "ihospital", ihospital())
                rv% = dce_push_array(Socket, "iareano", iareano())
                dce_table = dce_submit("sl_p_m05", "sl_areano_result_ul_4", Socket)
        End If
        sl_areano_result_ul_4 = dce_pop_long(dce_table, "dce_result")
        rv% = dce_pop_string(dce_table, "oerrmsg", oerrmsg)
        Call dce_table_destroy(dce_table)
End Function


Function sl_vit_download&(oerrmsg$, ispcid$())
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_p_m05")
        If (Socket > -1) Then
                rv% = dce_push_array(Socket, "ispcid", ispcid())
                dce_table = dce_submit("sl_p_m05", "sl_vit_download", Socket)
        End If
        sl_vit_download = dce_pop_long(dce_table, "dce_result")
        rv% = dce_pop_string(dce_table, "oerrmsg", oerrmsg)
        Call dce_table_destroy(dce_table)
End Function


'** ispcid          년월일 시분초,
'** iexamcode       검사코드,
'** iresult         결과값,
'** ierrflag        장비에서 나오는 오류값,
'** iequipcd        장비코드,
'** igubun           구분값 ---- 별 의미없음
Function sl_online_result_control&(oerrmsg$, ispcid$(), iexamcode$(), iresult$(), ierrflag$(), iequipcd$(), igubun$)
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_p_m05")
    If (Socket > -1) Then
        rv% = dce_push_array(Socket, "ispcid", ispcid())
        rv% = dce_push_array(Socket, "iexamcode", iexamcode())
        rv% = dce_push_array(Socket, "iresult", iresult())
        rv% = dce_push_array(Socket, "ierrflag", ierrflag())
        rv% = dce_push_array(Socket, "iequipcd", iequipcd())
        rv% = dce_push_string(Socket, "igubun", igubun)
        dce_table = dce_submit("sl_p_m05", "sl_online_result_control", Socket)
    End If
    sl_online_result_control = dce_pop_long(dce_table, "dce_result")
    rv% = dce_pop_string(dce_table, "oerrmsg", oerrmsg)
    Call dce_table_destroy(dce_table)
End Function

'Function sl_online_result_control&(oerrmsg$, ispcid$(), iexamcode$(), iresult$(), ierrflag$(), iequipcd$(), igubun$)
'    Dim dce_table As Long, Socket As Integer
'
'    Call dce_checkver(2, 0)
'    Socket = dce_findserver("sl_p_m05")
'    If (Socket > -1) Then
'        rv% = dce_push_array(Socket, "ispcid", ispcid())
'        rv% = dce_push_array(Socket, "iexamcode", iexamcode())
'        rv% = dce_push_array(Socket, "iresult", iresult())
'        rv% = dce_push_array(Socket, "ierrflag", ierrflag())
'        rv% = dce_push_array(Socket, "iequipcd", iequipcd())
'        rv% = dce_push_string(Socket, "igubun", igubun)
'        dce_table = dce_submit("sl_p_m05", "sl_online_result_control", Socket)
'    End If
'    sl_online_result_control = dce_pop_long(dce_table, "dce_result")
'    rv% = dce_pop_string(dce_table, "oerrmsg", oerrmsg)
'    Call dce_table_destroy(dce_table)
'End Function
