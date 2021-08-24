Attribute VB_Name = "sl_p_95_c"
Function sl_online_result_ul_r&(oerrmsg$, ispcid$(), iexamcode$(), iresult$(), ierrflag$(), iequipcd$(), igubun$)
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_p_95")
    If (Socket > -1) Then
        rv& = dce_push_array(Socket, "ispcid", ispcid())
        rv& = dce_push_array(Socket, "iexamcode", iexamcode())
        rv& = dce_push_array(Socket, "iresult", iresult())
        rv& = dce_push_array(Socket, "ierrflag", ierrflag())
        rv& = dce_push_array(Socket, "iequipcd", iequipcd())
        rv& = dce_push_string(Socket, "igubun", igubun)
        dce_table = dce_submit("sl_p_95", "sl_online_result_ul_r", Socket)
    End If
    sl_online_result_ul_r = dce_pop_long(dce_table, "dce_result")
    rv& = dce_pop_string(dce_table, "oerrmsg", oerrmsg)
    Call dce_table_destroy(dce_table)
End Function


Function sl_areano_result_ul_r&(oerrmsg$, iexamcode$(), iresult$(), ierrflag$(), iequipcd$(), iacptdate$(), islipcode$(), ihospital$(), iareano$())
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_p_95")
    If (Socket > -1) Then
        rv& = dce_push_array(Socket, "iexamcode", iexamcode())
        rv& = dce_push_array(Socket, "iresult", iresult())
        rv& = dce_push_array(Socket, "ierrflag", ierrflag())
        rv& = dce_push_array(Socket, "iequipcd", iequipcd())
        rv& = dce_push_array(Socket, "iacptdate", iacptdate())
        rv& = dce_push_array(Socket, "islipcode", islipcode())
        rv& = dce_push_array(Socket, "ihospital", ihospital())
        rv& = dce_push_array(Socket, "iareano", iareano())
        dce_table = dce_submit("sl_p_95", "sl_areano_result_ul_r", Socket)
    End If
    sl_areano_result_ul_r = dce_pop_long(dce_table, "dce_result")
    rv& = dce_pop_string(dce_table, "oerrmsg", oerrmsg)
    Call dce_table_destroy(dce_table)
End Function


Function sl_nr_result_ul&(oerrmsg$, ispcid$(), iexamcode$(), iresult$(), iequipcd$(), ierrflag$())
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_p_95")
    If (Socket > -1) Then
        rv& = dce_push_array(Socket, "ispcid", ispcid())
        rv& = dce_push_array(Socket, "iexamcode", iexamcode())
        rv& = dce_push_array(Socket, "iresult", iresult())
        rv& = dce_push_array(Socket, "iequipcd", iequipcd())
        rv& = dce_push_array(Socket, "ierrflag", ierrflag())
        dce_table = dce_submit("sl_p_95", "sl_nr_result_ul", Socket)
    End If
    sl_nr_result_ul = dce_pop_long(dce_table, "dce_result")
    rv& = dce_pop_string(dce_table, "oerrmsg", oerrmsg)
    Call dce_table_destroy(dce_table)
End Function


Function sl_online_barcode_ul_r&(oerrmsg$, ispcid$(), iexamcode$(), iresult$(), ierrflag$(), iequipcd$(), igubun$)
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_p_95")
    If (Socket > -1) Then
        rv& = dce_push_array(Socket, "ispcid", ispcid())
        rv& = dce_push_array(Socket, "iexamcode", iexamcode())
        rv& = dce_push_array(Socket, "iresult", iresult())
        rv& = dce_push_array(Socket, "ierrflag", ierrflag())
        rv& = dce_push_array(Socket, "iequipcd", iequipcd())
        rv& = dce_push_string(Socket, "igubun", igubun)
        dce_table = dce_submit("sl_p_95", "sl_online_barcode_ul_r", Socket)
    End If
    sl_online_barcode_ul_r = dce_pop_long(dce_table, "dce_result")
    rv& = dce_pop_string(dce_table, "oerrmsg", oerrmsg)
    Call dce_table_destroy(dce_table)
End Function


Function sl_p_95_p095_barcode_rslt&(oerrmsg$, ispcid$(), iexamcode$(), iresult$(), ierrflag$(), iequipcd$(), iuser_id$(), igubun$)
    Dim dce_table As Long, Socket As Integer

    Call dce_checkver(2, 0)
    Socket = dce_findserver("sl_p_95")
    If (Socket > -1) Then
        rv& = dce_push_array(Socket, "ispcid", ispcid())
        rv& = dce_push_array(Socket, "iexamcode", iexamcode())
        rv& = dce_push_array(Socket, "iresult", iresult())
        rv& = dce_push_array(Socket, "ierrflag", ierrflag())
        rv& = dce_push_array(Socket, "iequipcd", iequipcd())
        rv& = dce_push_array(Socket, "iuser_id", iuser_id())
        rv& = dce_push_string(Socket, "igubun", igubun)
        dce_table = dce_submit("sl_p_95", "sl_p_95_p095_barcode_rslt", Socket)
    End If
    sl_p_95_p095_barcode_rslt = dce_pop_long(dce_table, "dce_result")
    rv& = dce_pop_string(dce_table, "oerrmsg", oerrmsg)
    Call dce_table_destroy(dce_table)
End Function



