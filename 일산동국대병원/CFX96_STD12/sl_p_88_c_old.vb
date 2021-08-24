Function sl_online_result_ul_4& (oerrmsg$,ispcid$(),iexamcode$(),iresult$(),ierrflag$(),iequipcd$(),igubun$)
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_p_88")
	If (Socket > -1) Then
		rv% = dce_push_array(Socket,"ispcid",ispcid())
		rv% = dce_push_array(Socket,"iexamcode",iexamcode())
		rv% = dce_push_array(Socket,"iresult",iresult())
		rv% = dce_push_array(Socket,"ierrflag",ierrflag())
		rv% = dce_push_array(Socket,"iequipcd",iequipcd())
		rv% = dce_push_string(Socket,"igubun",igubun)
		dce_table = dce_submit("sl_p_88","sl_online_result_ul_4",Socket)
	End If
	sl_online_result_ul_4 = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_string(dce_table,"oerrmsg",oerrmsg)
	Call dce_table_destroy(dce_table)
End function


Function sl_areano_result_ul_4& (oerrmsg$,iexamcode$(),iresult$(),ierrflag$(),iequipcd$(),iacptdate$(),islipcode$(),ihospital$(),iareano$())
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_p_88")
	If (Socket > -1) Then
		rv% = dce_push_array(Socket,"iexamcode",iexamcode())
		rv% = dce_push_array(Socket,"iresult",iresult())
		rv% = dce_push_array(Socket,"ierrflag",ierrflag())
		rv% = dce_push_array(Socket,"iequipcd",iequipcd())
		rv% = dce_push_array(Socket,"iacptdate",iacptdate())
		rv% = dce_push_array(Socket,"islipcode",islipcode())
		rv% = dce_push_array(Socket,"ihospital",ihospital())
		rv% = dce_push_array(Socket,"iareano",iareano())
		dce_table = dce_submit("sl_p_88","sl_areano_result_ul_4",Socket)
	End If
	sl_areano_result_ul_4 = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_string(dce_table,"oerrmsg",oerrmsg)
	Call dce_table_destroy(dce_table)
End function


Function sl_vit_download& (oerrmsg$,ispcid$())
	dim dce_table as long, Socket as integer

	call dce_checkver(2,0)
	Socket = dce_findserver("sl_p_88")
	If (Socket > -1) Then
		rv% = dce_push_array(Socket,"ispcid",ispcid())
		dce_table = dce_submit("sl_p_88","sl_vit_download",Socket)
	End If
	sl_vit_download = dce_pop_long(dce_table,"dce_result")
	rv% = dce_pop_string(dce_table,"oerrmsg",oerrmsg)
	Call dce_table_destroy(dce_table)
End function


