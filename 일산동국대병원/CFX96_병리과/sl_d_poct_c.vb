Function sql_prepare_sl_d_poct& (db$,login$,pwd$)
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_poct")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"db",db)
		rv% = dce_push_string(Socket,"login",login)
		rv% = dce_push_string(Socket,"pwd",pwd)
		dce_table = dce_submit("sl_d_poct","sql_prepare_sl_d_poct",Socket)
	End If
	sql_prepare_sl_d_poct = dce_pop_long(dce_table,"dce_result")
	Call dce_table_destroy(dce_table)
End function


Function sql_rows_sl_d_poct& (maxrows&)
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_poct")
	If (Socket > -1) Then
		Call dce_push_long(Socket,"maxrows",maxrows)
		dce_table = dce_submit("sl_d_poct","sql_rows_sl_d_poct",Socket)
	End If
	sql_rows_sl_d_poct = dce_pop_long(dce_table,"dce_result")
	Call dce_table_destroy(dce_table)
End function


Function sql_set_max_rows_sl_d_poct& (maxrows&)
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_poct")
	If (Socket > -1) Then
		Call dce_push_long(Socket,"maxrows",maxrows)
		dce_table = dce_submit("sl_d_poct","sql_set_max_rows_sl_d_poct",Socket)
	End If
	sql_set_max_rows_sl_d_poct = dce_pop_long(dce_table,"dce_result")
	Call dce_table_destroy(dce_table)
End function


Function sl_d_poct_apemgrct& (vptno$,PT_NO$(),MEDDEPT$(),CHADR$(),PATSECT$())
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_poct")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"vptno",vptno)
		dce_table = dce_submit("sl_d_poct","sl_d_poct_apemgrct",Socket)
	End If
	sl_d_poct_apemgrct = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"PT_NO",PT_NO())
	rv% = dce_pop_array(dce_table,"MEDDEPT",MEDDEPT())
	rv% = dce_pop_array(dce_table,"CHADR",CHADR())
	rv% = dce_pop_array(dce_table,"PATSECT",PATSECT())
	Call dce_table_destroy(dce_table)
End function


Function sl_d_poct_apiplist& (vptno$,PT_NO$(),MEDDEPT$(),CHADR$(),PATSECT$())
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_poct")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"vptno",vptno)
		dce_table = dce_submit("sl_d_poct","sl_d_poct_apiplist",Socket)
	End If
	sl_d_poct_apiplist = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"PT_NO",PT_NO())
	rv% = dce_pop_array(dce_table,"MEDDEPT",MEDDEPT())
	rv% = dce_pop_array(dce_table,"CHADR",CHADR())
	rv% = dce_pop_array(dce_table,"PATSECT",PATSECT())
	Call dce_table_destroy(dce_table)
End function


Function sl_d_poct_moporeqt& (vptno$,PT_NO$(),MEDDEPT$(),CHADR$(),PATSECT$())
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_d_poct")
	If (Socket > -1) Then
		rv% = dce_push_string(Socket,"vptno",vptno)
		dce_table = dce_submit("sl_d_poct","sl_d_poct_moporeqt",Socket)
	End If
	sl_d_poct_moporeqt = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_array(dce_table,"PT_NO",PT_NO())
	rv% = dce_pop_array(dce_table,"MEDDEPT",MEDDEPT())
	rv% = dce_pop_array(dce_table,"CHADR",CHADR())
	rv% = dce_pop_array(dce_table,"PATSECT",PATSECT())
	Call dce_table_destroy(dce_table)
End function


