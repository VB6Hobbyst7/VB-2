Attribute VB_Name = "sl_d_m05"
Function sql_prepare_sl_d_m05&(db$, login$, pwd$)
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_d_m05")
        If (Socket > -1) Then
                rv% = dce_push_string(Socket, "db", db)
                rv% = dce_push_string(Socket, "login", login)
                rv% = dce_push_string(Socket, "pwd", pwd)
                dce_table = dce_submit("sl_d_m05", "sql_prepare_sl_d_m05", Socket)
        End If
        sql_prepare_sl_d_m05 = dce_pop_long(dce_table, "dce_result")
        Call dce_table_destroy(dce_table)
End Function


Function sql_rows_sl_d_m05&(maxrows&)
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_d_m05")
        If (Socket > -1) Then
                Call dce_push_long(Socket, "maxrows", maxrows)
                dce_table = dce_submit("sl_d_m05", "sql_rows_sl_d_m05", Socket)
        End If
        sql_rows_sl_d_m05 = dce_pop_long(dce_table, "dce_result")
        Call dce_table_destroy(dce_table)
End Function


Function sql_set_max_rows_sl_d_m05&(maxrows&)
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_d_m05")
        If (Socket > -1) Then
                Call dce_push_long(Socket, "maxrows", maxrows)
                dce_table = dce_submit("sl_d_m05", "sql_set_max_rows_sl_d_m05", Socket)
        End If
        sql_set_max_rows_sl_d_m05 = dce_pop_long(dce_table, "dce_result")
        Call dce_table_destroy(dce_table)
End Function

'-- 오더조회
Function sl_spcid_tstcd_select&(SpcNo$, tst_cd$())
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_d_m05")
        If (Socket > -1) Then
                rv% = dce_push_string(Socket, "spcno", SpcNo)
                dce_table = dce_submit("sl_d_m05", "sl_spcid_tstcd_select", Socket)
       End If
        sl_spcid_tstcd_select = dce_pop_long(dce_table, "dce_result")
        rv% = dce_pop_array(dce_table, "tst_cd", tst_cd())
        Call dce_table_destroy(dce_table)
End Function


Function sel_order_Total_select&(spc_no$(), tst_cd$())
        Dim dce_table As Long, Socket As Integer

        Call dce_checkver(2, 0)
        Socket = dce_findserver("sl_d_m05")
        If (Socket > -1) Then
                dce_table = dce_submit("sl_d_m05", "sel_order_Total_select", Socket)
        End If
        sel_order_Total_select = dce_pop_long(dce_table, "dce_result")
        rv% = dce_pop_array(dce_table, "spc_no", spc_no())
        rv% = dce_pop_array(dce_table, "tst_cd", tst_cd())
        Call dce_table_destroy(dce_table)
End Function


