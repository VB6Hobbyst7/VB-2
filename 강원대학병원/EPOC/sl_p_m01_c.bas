Attribute VB_Name = "Module1"
Option Explicit

Function sl_online_pc_vitck&(oerrmsg$, i_pt_no$, i_acptdte$, i_acptno$, i_isotcd$, i_isotrlst$, i_pn$, i_maccd$, i_viteckid$, i_anticd$(), i_diam$(), i_rslt$(), i_antinm$(), count_in&)
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_p_m01")
        If (Socket > -1) Then
                rv% = dce_push_string(Socket, "i_pt_no", i_pt_no)
                rv% = dce_push_string(Socket, "i_acptdte", i_acptdte)
                rv% = dce_push_string(Socket, "i_acptno", i_acptno)
                rv% = dce_push_string(Socket, "i_isotcd", i_isotcd)
                rv% = dce_push_string(Socket, "i_isotrlst", i_isotrlst)
                rv% = dce_push_string(Socket, "i_pn", i_pn)
                rv% = dce_push_string(Socket, "i_maccd", i_maccd)
                rv% = dce_push_string(Socket, "i_viteckid", i_viteckid)
                rv% = dce_push_array(Socket, "i_anticd", i_anticd())
                rv% = dce_push_array(Socket, "i_diam", i_diam())
                rv% = dce_push_array(Socket, "i_rslt", i_rslt())
                rv% = dce_push_array(Socket, "i_antinm", i_antinm())
                Call dce_push_long(Socket, "count_in", count_in)
                dce_table = dce_submit("sl_p_m01", "sl_online_pc_vitck", Socket)
        End If
        sl_online_pc_vitck = dce_pop_long(dce_table, "dce_result")
        rv% = dce_pop_string(dce_table, "oerrmsg", oerrmsg)
        Call dce_table_destroy(dce_table)
End Function


Function sl_online_result_control&(oerrmsg$, ispcid$(), iexamcode$(), iresult$(), ierrflag$(), iequipcd$(), igubun$)
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_p_m01")
        If (Socket > -1) Then
                rv% = dce_push_array(Socket, "ispcid", ispcid())
                rv% = dce_push_array(Socket, "iexamcode", iexamcode())
                rv% = dce_push_array(Socket, "iresult", iresult())
                rv% = dce_push_array(Socket, "ierrflag", ierrflag())
                rv% = dce_push_array(Socket, "iequipcd", iequipcd())
                rv% = dce_push_string(Socket, "igubun", igubun)
                dce_table = dce_submit("sl_p_m01", "sl_online_result_control", Socket)
        End If
        sl_online_result_control = dce_pop_long(dce_table, "dce_result")
        rv% = dce_pop_string(dce_table, "oerrmsg", oerrmsg)
        Call dce_table_destroy(dce_table)
End Function


Function sl_online_pc_98&(oerrmsg$, ispcid$(), iexamcode$(), iresult$(), ierrflag$(), iequipcd$(), igubun$)
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_p_m01")
        If (Socket > -1) Then
                rv% = dce_push_array(Socket, "ispcid", ispcid())
                rv% = dce_push_array(Socket, "iexamcode", iexamcode())
                rv% = dce_push_array(Socket, "iresult", iresult())
                rv% = dce_push_array(Socket, "ierrflag", ierrflag())
                rv% = dce_push_array(Socket, "iequipcd", iequipcd())
                rv% = dce_push_string(Socket, "igubun", igubun)
                dce_table = dce_submit("sl_p_m01", "sl_online_pc_98", Socket)
        End If
        sl_online_pc_98 = dce_pop_long(dce_table, "dce_result")
        rv% = dce_pop_string(dce_table, "oerrmsg", oerrmsg)
        Call dce_table_destroy(dce_table)
End Function


