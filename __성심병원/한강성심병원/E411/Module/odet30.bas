Attribute VB_Name = "Module2"
Option Explicit

'******************** Declare ODE LITE  API  Functions ************************
'******************************************************************************

#If Win32 Then
Declare Function dce_setenv& Lib "odet30.dll" (ByVal s1$, ByVal s2$, ByVal s3$)
#Else
Declare Function dce_setenv% Lib "odet30.dll" (ByVal s1$, ByVal s2$, ByVal s3$)
#End If
#If Win32 Then
Declare Function dce_error& Lib "odet30.dll" (ByVal s1$)
#Else
Declare Function dce_error% Lib "odet30.dll" (ByVal s1$)
#End If
#If Win32 Then
Declare Sub dce_seterr Lib "odet30.dll" (ByVal errnum&, ByVal errstr$)
#Else
Declare Sub dce_seterr Lib "odet30.dll" (ByVal errnum%, ByVal errstr$)
#End If
Declare Sub dce_log_str Lib "odet30.dll" (ByVal MSG$)
#If Win32 Then
Declare Function dce_err_is_fatal& Lib "odet30.dll" ()
#Else
Declare Function dce_err_is_fatal% Lib "odet30.dll" ()
#End If
#If Win32 Then
Declare Function dce_set_account& Lib "odet30.dll" (ByVal s1$, ByVal s2$)
#Else
Declare Function dce_set_account% Lib "odet30.dll" (ByVal s1$, ByVal s2$)
#End If
#If Win32 Then
Declare Function dce_errnum& Lib "odet30.dll" ()
#Else
Declare Function dce_errnum% Lib "odet30.dll" ()
#End If
#If Win32 Then
Declare Function dce_ignore_input& Lib "odet30.dll" (ByVal mode&)
#Else
Declare Function dce_ignore_input% Lib "odet30.dll" (ByVal mode%)
#End If
#If Win32 Then
Declare Function dce_close_env& Lib "odet30.dll" ()
#Else
Declare Function dce_close_env% Lib "odet30.dll" ()
#End If
Declare Sub dce_clearerror Lib "odet30.lib" ()
Declare Function dce_get_avg_idle& Lib "odet30.lib" (ByVal aHost$, ByVal aPort$)

#If Win32 Then
Declare Function dce_submit& Lib "odet30.dll" (ByVal s1$, ByVal s2$, ByVal Socket&)
#Else
Declare Function dce_submit& Lib "odet30.dll" (ByVal s1$, ByVal s2$, ByVal Socket%)
#End If
#If Win32 Then
Declare Function dce_findserver& Lib "odet30.dll" (ByVal s1$)
#Else
Declare Function dce_findserver% Lib "odet30.dll" (ByVal s1$)
#End If
Declare Function dce_get_len& Lib "odet30.dll" (ByVal ptr As Any)
Declare Function dce_get_width& Lib "odet30.dll" (ByVal ptr As Any)
Declare Function dce_get_rows& Lib "odet30.dll" (ByVal ptr As Any)
Declare Function dce_element_add& Lib "odet30.dll" (ByVal ptr As Any, ByVal s1$)
#If Win32 Then
Declare Function dce_send_array& Lib "odet30.dll" (ByVal n1&, ByVal s1$, ByVal ptr As Any)
#Else
Declare Function dce_send_array% Lib "odet30.dll" (ByVal n1%, ByVal s1$, ByVal ptr As Any)
#End If
Declare Function dce_element_get& Lib "odet30.dll" (ByVal ptr As Any, ByVal s1$, ByVal s2$, ByVal n1&)

'***************************************************************
'THESE FUNCTIONS WILL NOT WORK PROPERLY, BECAUSE YOU CANNOT PASS
'ARRAYS OF STRINGS
'Declare Function dce_get_array_info& Lib "odet30.dll" (ByVal ptr As Any, ByVal s1$, ByVal n1&, ByVal n2&)
'Declare Function dce_get_array& Lib "odet30.dll" (ByVal ptr As Any, ByVal strptr As Any)
'***************************************************************

Declare Function dce_table_find& Lib "odet30.dll" (ByVal ptr As Any, ByVal s1$)
Declare Sub dce_table_destroy Lib "odet30.dll" (ByVal ptr As Any)
#If Win32 Then
Declare Function dce_table_push& Lib "odet30.dll" (ByVal Socket&, ByVal s1$, ByVal s2$, ByVal n1&)
#Else
Declare Function dce_table_push% Lib "odet30.dll" (ByVal Socket%, ByVal s1$, ByVal s2$, ByVal n1&)
#End If
#If Win32 Then
Declare Function dce_table_pop& Lib "odet30.dll" (ByVal ptr As Any, ByVal s1$, ByVal s2$, ByVal n1&)
#Else
Declare Function dce_table_pop% Lib "odet30.dll" (ByVal ptr As Any, ByVal s1$, ByVal s2$, ByVal n1&)
#End If

#If Win32 Then
Declare Sub dce_push_int Lib "odet30.dll" (ByVal Socket&, ByVal s1$, ByVal num&)
#Else
Declare Sub dce_push_int Lib "odet30.dll" (ByVal Socket%, ByVal s1$, ByVal num%)
#End If
#If Win32 Then
Declare Sub dce_push_long Lib "odet30.dll" (ByVal Socket&, ByVal s1$, ByVal num&)
#Else
Declare Sub dce_push_long Lib "odet30.dll" (ByVal Socket%, ByVal s1$, ByVal num&)
#End If
#If Win32 Then
Declare Sub dce_push_short Lib "odet30.dll" (ByVal Socket&, ByVal s1$, ByVal num%)
#Else
Declare Sub dce_push_short Lib "odet30.dll" (ByVal Socket%, ByVal s1$, ByVal num%)
#End If

#If Win32 Then
Declare Function dce_pop_int& Lib "odet30.dll" (ByVal ptr As Any, ByVal s1$)
#Else
Declare Function dce_pop_int% Lib "odet30.dll" (ByVal ptr As Any, ByVal s1$)
#End If
Declare Function dce_pop_long& Lib "odet30.dll" (ByVal ptr As Any, ByVal s1$)
Declare Function dce_pop_short% Lib "odet30.dll" (ByVal ptr As Any, ByVal s1$)

Declare Sub dce_release Lib "odet30.dll" ()
Declare Sub dce_version Lib "odet30.dll" (ByVal s1$)

'**************************************************************************
' Dedicated Server functions
'**************************************************************************
Declare Sub dce_dedDisconnect Lib "odet30.dll" (ByVal aServName$)
#If Win32 Then
Declare Function dce_isDedicated& Lib "odet30.dll" (ByVal aServName$)
#Else
Declare Function dce_isDedicated% Lib "odet30.dll" (ByVal aServName$)
#End If

'**************************************************************************
' Version compatibility function
'**************************************************************************
#If Win32 Then
Declare Sub dce_checkver Lib "odet30.dll" (ByVal major&, ByVal minor&)
#Else
Declare Sub dce_checkver Lib "odet30.dll" (ByVal major%, ByVal minor%)
#End If


'**************************************************************************

Sub arr2list(myarray() As String, mylist As Control)
    Dim i%, j%
    For i = 0 To mylist.ListCount - 1
        mylist.RemoveItem 0
    Next i

    j = 0
    For i = LBound(myarray) To UBound(myarray)
        mylist.List(j) = myarray(i)
        j = j + 1
    Next i
End Sub

Sub ClearObject(Object As Control)
    Dim i%
    If TypeOf Object Is ListBox Then
        For i = 0 To Object.ListCount - 1
            Object.RemoveItem 0
        Next i
    ElseIf TypeOf Object Is PictureBox Then
        Object.BackColor = &H80000005
    ElseIf TypeOf Object Is TextBox Then
        Object.text = ""
    ElseIf TypeOf Object Is label Then
        Object.Caption = ""
    End If
End Sub

Sub dce_c2vbdouble_str(double_str$)
'This function converts ###e### to ###D###

Dim exp_marker_pos%
exp_marker_pos = InStr(double_str, "e")
If exp_marker_pos > 0 Then
    double_str = Mid$(double_str, 1, exp_marker_pos - 1) + "D" + Mid$(double_str, exp_marker_pos + 1)
End If

End Sub

Function dce_iserror%()
    Dim ErrMsg As String * 250
    dce_iserror = dce_error(ErrMsg)
End Function

Function dce_pop_array&(ptr&, tag$, mydata() As String)
    Dim nptr&, counter&, myrows&, mywidth&, TheData$, idx&
    Dim rc%
    Dim NONULL&
    Dim ANULL$
    ANULL = Chr$(0)

    'First check the library error flag
    If (dce_iserror() <> 0) Then
        dce_pop_array = -1
        Exit Function
    End If
    
    nptr = dce_table_find(ptr, tag)
    myrows = 0
    If (nptr <> 0) Then
        myrows = dce_get_rows(nptr)
        If (myrows <> 0) Then
          ReDim mydata(myrows)
          counter = LBound(mydata)
          mywidth = dce_get_width(nptr)
          idx = 0
        Do While (1 = 1)
            TheData = String$(mywidth + 1, 0)
            idx = dce_element_get(nptr, tag, TheData, idx)
            If (idx = 0) Then
                Exit Do
            End If
            NONULL = InStr(TheData, ANULL)
            If (NONULL <> 0) Then
                mydata(counter) = left$(TheData, NONULL - 1)
            Else
                mydata(counter) = ""
            End If
            counter = counter + 1
        Loop
        dce_pop_array = myrows
       Else
        ReDim mydata(0)
        dce_pop_array = 0
       End If
    Else
        'The data was not found in the dce_table
        ReDim mydata(0)
        dce_pop_array = 0
    End If

End Function

Function dce_pop_char$(dce_table&, tag$)
    Dim nptr&, mylen&, mydata$, idx&

    'First check the library error flag
    If (dce_iserror() <> 0) Then
        dce_pop_char = ""
        Exit Function
    End If

    nptr = dce_table_find(dce_table, tag)
    If (nptr <> 0) Then
        mylen = dce_get_len(nptr)
        mydata = String$(mylen, 0)
        idx = dce_element_get(nptr, tag, mydata, 0)
        dce_pop_char = left$(mydata, 1)
    Else
        'The data was not found in the dce_table
        dce_pop_char = ""
    End If

End Function

Sub dce_pop_char_arr(dce_table&, tag$, mydata() As String, myrow&, mycol&)
    Dim nptr&, counter&, myrows&, mywidth&, TheData$, idx&
    Dim maxarr&, NONULL&, ANULL$
    ANULL = Chr$(0)

    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If

    If (myrow > UBound(mydata) - LBound(mydata) + 1) Then
        'Not enough space has been allocated
        TheData = "<OEC VB ERROR>  dce_pop_char_arr: <" + tag + "> array size too small"
        Call dce_log_str(TheData)
        Call dce_seterr(34, "dce_pop_char_arr")
        Exit Sub
    End If

    nptr = dce_table_find(dce_table, tag)
    myrows = 0
    counter = LBound(mydata)

    If (nptr <> 0) Then
        myrows = dce_get_rows(nptr)

        If (myrow < myrows) Then
            maxarr = counter + myrow
        Else
            maxarr = counter + myrows
        End If

        mywidth = dce_get_width(nptr)
        idx = 0
        
        Do While (counter < maxarr)
            TheData = String$(mywidth + 1, 0)
            idx = dce_element_get(nptr, tag, TheData, idx)
            If (idx = 0) Then
                Exit Do
            End If

            NONULL = InStr(TheData, ANULL)
            If NONULL > mycol Then
                NONULL = mycol
            End If

            If (NONULL <> 0) Then
                mydata(counter) = left$(TheData, NONULL - 1)
            Else
                mydata(counter) = ""
            End If
            counter = counter + 1
        Loop
    End If

    'Now fill the rest of the elements with empty strings
    maxarr = LBound(mydata) + myrow
    Do While (counter < maxarr)
        mydata(counter) = ""
        counter = counter + 1
    Loop

End Sub

Sub dce_pop_char_Darr(dce_table&, tag$, mydata() As String, myrow&, mycol&)
    Dim nptr&, counter&, myrows&, mywidth&, TheData$, idx&
    nptr = dce_table_find(dce_table, tag)
    Dim NONULL&, ANULL$
    ANULL = Chr$(0)

    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If
    
    myrows = 0
    If (nptr <> 0) Then
        myrows = dce_get_rows(nptr)
        If (myrows <> myrow) Then 'The stub is bad
            TheData = "<OEC VB ERROR>  dce_pop_char_Darr <" + tag + "> wrong rowsize:  expected <" + Str$(myrow) + "> got <" + Str$(myrows) + ">"
            Call dce_log_str(TheData)
            Call dce_seterr(34, "dce_pop_char_Darr")
            Exit Sub
        Else
            ReDim mydata(myrows)
            counter = LBound(mydata)
            mywidth = dce_get_width(nptr)
            If (mywidth + 1 > mycol) Then
                TheData = "<OEC VB ERROR> dce_pop_char_Darr <" + tag + "> column size too big:  expected <" + Str$(mycol) + "> got <" + Str$(mywidth + 1) + ">"
                Call dce_log_str(TheData)
                Call dce_seterr(34, "dce_pop_char_Darr")
                Exit Sub
            End If

            idx = 0
            Do While (1 = 1)
                TheData = String$(mywidth + 1, 0)
                idx = dce_element_get(nptr, tag, TheData, idx)
                If (idx = 0) Then
                    Exit Do
                End If
                NONULL = InStr(TheData, ANULL)
                If (NONULL <> 0) Then
                    mydata(counter) = left$(TheData, NONULL - 1)
                Else
                    mydata(counter) = ""
                End If
                counter = counter + 1
            Loop
        End If
    Else
        'No data found in dce_table:  no rows sent
        ReDim mydata(0)
    End If

End Sub

Sub dce_pop_char_Dstr(dce_table&, tag$, mydata$, mylen&)
    Dim nptr&, arr_len&, idx&, TheData$
    Dim ANULL$
    Dim NONULL&
    ANULL = Chr$(0)

    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If
    
    nptr = dce_table_find(dce_table, tag)
    If (nptr <> 0) Then
        arr_len = dce_get_len(nptr)
        If (mylen <> arr_len) Then
            TheData = "<OEC VB ERROR>  dce_pop_char_Dstr <" + tag + "> wrong length:  expected <" + Str$(mylen) + "> got <" + Str$(arr_len) + ">"
            Call dce_log_str(TheData)
            Call dce_seterr(34, "dce_pop_char_Dstr")
            Exit Sub
        End If

        TheData = String$(arr_len, 0)
        idx = dce_element_get(nptr, tag, TheData, 0)
        NONULL = InStr(TheData, ANULL)
        If (NONULL <> 0) Then
            mydata = left$(TheData, NONULL - 1)
        Else
            mydata = ""
        End If

    Else
        mydata = ""
    End If

End Sub

Sub dce_pop_char_str(dce_table&, tag$, mydata$, mylen&)
    Dim nptr&, arr_len&, idx&, TheData$
    Dim ANULL$
    ANULL = Chr$(0)

    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If
    
    If (mylen > Len(mydata) + 1) Then
        'Not enough space has been allocated
        TheData = "<OEC VB ERROR>  dce_pop_char_str <" + tag + "> array size passed in too small"
        Call dce_log_str(TheData)
        Call dce_seterr(34, "dce_pop_char_str")
        Exit Sub
    End If
    
    nptr = dce_table_find(dce_table, tag)
    If (nptr <> 0) Then

        arr_len = dce_get_len(nptr)

        TheData = String$(arr_len, 0)
        idx = dce_element_get(nptr, tag, TheData, 0)
        TheData = left$(TheData, mylen - 1) + ANULL
        mydata = left$(TheData, InStr(TheData, ANULL) - 1)
        'The rest of the memory is not touched
    Else
        mydata = ""
    End If


End Sub

Function dce_pop_double#(dce_table&, tag$)
    Dim nptr&, mylen&, mydata$, idx&

    'First check the library error flag
    If (dce_iserror() <> 0) Then
        dce_pop_double = 0
        Exit Function
    End If

    nptr = dce_table_find(dce_table, tag)
    If (nptr <> 0) Then
        mylen = dce_get_len(nptr)
        mydata = String$(mylen, 0)
        idx = dce_element_get(nptr, tag, mydata, 0)
        Call dce_c2vbdouble_str(mydata)
        dce_pop_double = Val(mydata)
    Else
        dce_pop_double = 0
    End If

End Function

Sub dce_pop_double_Dstr(dce_table&, tag$, mydata() As Double, myrow&)
    Dim nptr&, counter&, myrows&, mywidth&, TheData$, idx&
    Dim NONULL&, ANULL$
    ANULL = Chr$(0)

    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If

    nptr = dce_table_find(dce_table, tag)
    If (nptr <> 0) Then
        myrows = dce_get_rows(nptr&)
        If (myrows <> myrow) Then
            TheData = "<OEC VB ERROR>  dce_pop_double_Dstr <" + tag + "> wrong rowsize:  expected <" + Str$(myrow) + "> got <" + Str$(myrows) + ">"
            Call dce_log_str(TheData)
            Call dce_seterr(34, "dce_pop_double_Dstr")
            Exit Sub
        
        
        Else
            ReDim mydata(myrows)
            counter = LBound(mydata)
            mywidth = dce_get_width(nptr)
            TheData = String$(mywidth + 1, 0)
            idx = 0
            Do While (1 = 1)
                idx = dce_element_get(nptr, tag, TheData, idx)
                If (idx = 0) Then
                    Exit Do
                End If
                Call dce_c2vbdouble_str(TheData)
                mydata(counter) = Val(TheData)
                counter = counter + 1
            Loop
        End If
    Else
        'No data found in dce_table
        ReDim mydata(0)
    End If

End Sub

Sub dce_pop_double_str(dce_table&, tag$, mydata() As Double, myrow&)
    Dim nptr&, counter&, myrows&, mywidth&, TheData$, idx&
    Dim maxarr&

    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If

    If (myrow > UBound(mydata) - LBound(mydata) + 1) Then
        'Not enough space has been allocated
        TheData = "<OEC VB ERROR>  dce_pop_double_str: <" + tag + "> array size too small"
        Call dce_log_str(TheData)
        Call dce_seterr(34, "dce_pop_double_str")
        Exit Sub
    End If

    nptr = dce_table_find(dce_table, tag)
    counter = LBound(mydata)
    
    If (nptr <> 0) Then
        myrows = dce_get_rows(nptr)

        If (myrow < myrows) Then
            maxarr = counter + myrow
        Else
            maxarr = counter + myrows
        End If

        mywidth = dce_get_width(nptr)
        TheData = String$(mywidth + 1, 0)
        idx = 0
        Do While (counter < maxarr)
            idx = dce_element_get(nptr, tag, TheData, idx)
            If (idx = 0) Then
                Exit Do
            End If
            Call dce_c2vbdouble_str(TheData)
            mydata(counter) = Val(TheData)
            counter = counter + 1
        Loop
    End If

    'Now fill the rest of the elements with 0
    maxarr = LBound(mydata) + myrow
    Do While (counter < maxarr)
        mydata(counter) = 0
        counter = counter + 1
    Loop

End Sub

Function dce_pop_float!(dce_table&, tag$)
    Dim nptr&, mylen&, mydata$, idx&

    'First check the library error flag
    If (dce_iserror() <> 0) Then
        dce_pop_float = 0
        Exit Function
    End If

    nptr = dce_table_find(dce_table, tag)
    If (nptr <> 0) Then
        mylen = dce_get_len(nptr)
        mydata = String$(mylen, 0)
        idx = dce_element_get(nptr, tag, mydata, 0)
        dce_pop_float = Val(mydata)
    Else
        'The data was not found in the dce_table
        dce_pop_float = 0
    End If

End Function

Sub dce_pop_float_Dstr(dce_table&, tag$, mydata() As Single, myrow&)
    Dim nptr&, counter&, myrows&, mywidth&, TheData$, idx&
    Dim NONULL&, ANULL$
    ANULL = Chr$(0)

    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If

    nptr = dce_table_find(dce_table, tag)
    If (nptr <> 0) Then
        myrows = dce_get_rows(nptr&)
        If (myrows <> myrow) Then
            TheData = "<OEC VB ERROR>  dce_pop_float_Dstr <" + tag + "> wrong rowsize:  expected <" + Str$(myrow) + "> got <" + Str$(myrows) + ">"
            Call dce_log_str(TheData)
            Call dce_seterr(34, "dce_pop_float_Dstr")
            Exit Sub
        
        
        Else
            ReDim mydata(myrows)
            counter = LBound(mydata)
            mywidth = dce_get_width(nptr)
            TheData = String$(mywidth + 1, 0)
            idx = 0
            Do While (1 = 1)
                idx = dce_element_get(nptr, tag, TheData, idx)
                If (idx = 0) Then
                    Exit Do
                End If
                mydata(counter) = Val(TheData)
                counter = counter + 1
            Loop
        End If
    Else
        'No data found in dce_table
        ReDim mydata(0)
    End If

End Sub

Sub dce_pop_float_str(dce_table&, tag$, mydata() As Single, myrow&)
    Dim nptr&, counter&, myrows&, mywidth&, TheData$, idx&
    Dim maxarr&

    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If

    If (myrow > UBound(mydata) - LBound(mydata) + 1) Then
        'Not enough space has been allocated
        TheData = "<OEC VB ERROR>  dce_pop_float_str: <" + tag + "> array size too small"
        Call dce_log_str(TheData)
        Call dce_seterr(34, "dce_pop_float_str")
        Exit Sub
    End If

    nptr = dce_table_find(dce_table, tag)
    counter = LBound(mydata)

    If (nptr <> 0) Then
        myrows = dce_get_rows(nptr)

        If (myrow < myrows) Then
            maxarr = counter + myrow
        Else
            maxarr = counter + myrows
        End If


        mywidth = dce_get_width(nptr)
        TheData = String$(mywidth + 1, 0)
        idx = 0
        Do While (counter < maxarr)
            idx = dce_element_get(nptr, tag, TheData, idx)
            If (idx = 0) Then
                Exit Do
            End If
            mydata(counter) = Val(TheData)
            counter = counter + 1
        Loop
    End If

    'Now fill the rest of the elements with 0
    maxarr = LBound(mydata) + myrow
    Do While (counter < maxarr)
        mydata(counter) = 0
        counter = counter + 1
    Loop

End Sub

Sub dce_pop_int_Dstr(dce_table&, tag$, mydata() As Integer, myrow&)
    Dim nptr&, counter&, myrows&, mywidth&, TheData$, idx&
    Dim NONULL&, ANULL$
    ANULL = Chr$(0)

    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If

    nptr = dce_table_find(dce_table, tag)
    If (nptr <> 0) Then
        myrows = dce_get_rows(nptr&)
        If (myrows <> myrow) Then
            TheData = "<OEC VB ERROR>  dce_pop_int_Dstr <" + tag + "> wrong rowsize:  expected <" + Str$(myrow) + "> got <" + Str$(myrows) + ">"
            Call dce_log_str(TheData)
            Call dce_seterr(34, "dce_pop_int_Dstr")
            Exit Sub
        
        
        Else
            ReDim mydata(myrows)
            counter = LBound(mydata)
            mywidth = dce_get_width(nptr)
            TheData = String$(mywidth + 1, 0)
            idx = 0
            Do While (1 = 1)
                idx = dce_element_get(nptr, tag, TheData, idx)
                If (idx = 0) Then
                    Exit Do
                End If
                mydata(counter) = Val(TheData)
                counter = counter + 1
            Loop
        End If
    Else
        'No data found in dce_table
        ReDim mydata(0)
    End If

End Sub

Sub dce_pop_int_str(dce_table&, tag$, mydata() As Integer, myrow&)
    Dim nptr&, counter&, myrows&, mywidth&, TheData$, idx&
    Dim maxarr&

    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If

    If (myrow > UBound(mydata) - LBound(mydata) + 1) Then
        'Not enough space has been allocated
        TheData = "<OEC VB ERROR>  dce_pop_int_str: <" + tag + "> array size too small"
        Call dce_log_str(TheData)
        Call dce_seterr(34, "dce_pop_int_str")
        Exit Sub
    End If

    nptr = dce_table_find(dce_table, tag)
    counter = LBound(mydata)
    If (nptr <> 0) Then
        myrows = dce_get_rows(nptr)

        If (myrow < myrows) Then
            maxarr = counter + myrow
        Else
            maxarr = counter + myrows
        End If


        mywidth = dce_get_width(nptr)
        TheData = String$(mywidth + 1, 0)
        idx = 0
        Do While (counter < maxarr)
            idx = dce_element_get(nptr, tag, TheData, idx)
            If (idx = 0) Then
                Exit Do
            End If
            mydata(counter) = Val(TheData)
            counter = counter + 1
        Loop
    End If

    'Now fill the rest of the elements with 0
    maxarr = LBound(mydata) + myrow
    Do While (counter < maxarr)
        mydata(counter) = 0
        counter = counter + 1
    Loop

End Sub

Sub dce_pop_long_Dstr(dce_table&, tag$, mydata() As Long, myrow&)
    Dim nptr&, counter&, myrows&, mywidth&, TheData$, idx&
    Dim NONULL&, ANULL$
    ANULL = Chr$(0)

    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If

    nptr = dce_table_find(dce_table, tag)
    If (nptr <> 0) Then
        myrows = dce_get_rows(nptr&)
        If (myrows <> myrow) Then
            TheData = "<OEC VB ERROR>  dce_pop_long_Dstr <" + tag + "> wrong rowsize:  expected <" + Str$(myrow) + "> got <" + Str$(myrows) + ">"
            Call dce_log_str(TheData)
            Call dce_seterr(34, "dce_pop_long_Dstr")
            Exit Sub
        
        
        Else
            ReDim mydata(myrows)
            counter = LBound(mydata)
            mywidth = dce_get_width(nptr)
            TheData = String$(mywidth + 1, 0)
            idx = 0
            Do While (1 = 1)
                idx = dce_element_get(nptr, tag, TheData, idx)
                If (idx = 0) Then
                    Exit Do
                End If
                mydata(counter) = Val(TheData)
                counter = counter + 1
            Loop
        End If
    Else
        'No data found in dce_table
        ReDim mydata(0)
    End If
End Sub

Sub dce_pop_long_str(dce_table&, tag$, mydata() As Long, myrow&)
    Dim nptr&, counter&, myrows&, mywidth&, TheData$, idx&
    Dim maxarr&

    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If

    If (myrow > UBound(mydata) - LBound(mydata) + 1) Then
        'Not enough space has been allocated
        TheData = "<OEC VB ERROR>  dce_pop_long_str: <" + tag + "> array size too small"
        Call dce_log_str(TheData)
        Call dce_seterr(34, "dce_pop_long_str")
        Exit Sub
    End If

    nptr = dce_table_find(dce_table, tag)
    counter = LBound(mydata)
    If (nptr <> 0) Then
        myrows = dce_get_rows(nptr)

        If (myrow < myrows) Then
            maxarr = counter + myrow
        Else
            maxarr = counter + myrows
        End If


        mywidth = dce_get_width(nptr)
        TheData = String$(mywidth + 1, 0)
        idx = 0
        Do While (counter < maxarr)
            idx = dce_element_get(nptr, tag, TheData, idx)
            If (idx = 0) Then
                Exit Do
            End If
            mydata(counter) = Val(TheData)
            counter = counter + 1
        Loop
    End If

    'Now fill the rest of the elements with 0
    maxarr = LBound(mydata) + myrow
    Do While (counter < maxarr)
        mydata(counter) = 0
        counter = counter + 1
    Loop

End Sub

Sub dce_pop_short_Dstr(dce_table&, tag$, mydata() As Integer, mylen&)
Call dce_pop_int_Dstr(dce_table, tag, mydata(), mylen)
End Sub

Sub dce_pop_short_str(dce_table&, tag$, mydata() As Integer, mylen&)
Call dce_pop_int_str(dce_table, tag, mydata(), mylen)
End Sub

Function dce_pop_string&(dce_table&, tag$, mydata$)
    Dim nptr&, mylen&, idx&, TheData$, ANULL$, NONULL&
    ANULL = Chr$(0)
    
    'First check the library error flag
    If (dce_iserror() <> 0) Then
        dce_pop_string = -1
        Exit Function
    End If
    
    nptr = dce_table_find(dce_table, tag)
    If (nptr <> 0) Then
        mylen = dce_get_len(nptr)
        TheData = String$(mylen, 0)
        idx = dce_element_get(nptr, tag, TheData, 0)
        NONULL = InStr(TheData, ANULL)
        If (NONULL <> 0) Then
            mydata = left$(TheData, NONULL - 1)
        Else
            mydata = ""
        End If
        dce_pop_string = mylen
    Else
        'The data was not found in the dce_table
        mydata = ""
        dce_pop_string = 0
    End If

End Function

Sub dce_pop_void_Dstr(dce_table&, tag$, mydata$, mylen&)
    Dim rc%, arr_len&, TheData$, nptr&
    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If
    

    mydata = String$(mylen, 0)
    rc = dce_table_pop(dce_table, tag, mydata, mylen)
End Sub

Sub dce_pop_void_str(dce_table&, tag$, mydata$, mylen&)
    Dim rc%, arr_len&, TheData$, nptr&
    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If
    
    If (mylen > Len(mydata)) Then
        TheData = "<OEC VB ERROR>  dce_pop_void_str <" + tag + "> wrong input size passed in:  expected <" + Str$(mylen) + "> got <" + Str$(Len(mydata)) + ">"
        Call dce_log_str(TheData)
        Call dce_seterr(34, "dce_pop_void_str")
        Exit Sub
    End If


    rc = dce_table_pop(dce_table, tag, mydata, mylen)
    
End Sub

Function dce_push_array&(Socket%, tag$, mydata() As String)
    
    Dim ptr&, counter&, maxarr&, rv%, TheData$
    
    'First check the library error flag
    If (dce_iserror() <> 0) Then
        dce_push_array = -1
        Exit Function
    End If
    
    ptr = 0
    counter = LBound(mydata)
    maxarr = UBound(mydata) + 1
    While (counter < maxarr)
        TheData = mydata(counter) + Chr$(0)
        ptr = dce_element_add(ptr, TheData)
        counter = counter + 1
    Wend
    rv = dce_send_array(Socket, tag, ptr)
    dce_push_array = counter
End Function

Sub dce_push_char(Socket%, tag$, letter$)
    Dim char_str$, rc%

    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If

    char_str = left$(letter, 1) + Chr$(0)
    rc = dce_table_push(Socket, tag, char_str, 2)
End Sub

Sub dce_push_char_arr(Socket%, tag$, mydata() As String, myrow&, mycol&)
    Dim ptr&, counter&, maxarr&, TheData$, mysize&, rv%
    
    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If
    
    ptr = 0
    counter = LBound(mydata)
    mysize = UBound(mydata) - counter + 1
    If (mysize < myrow) Then
        maxarr = mysize + counter
    Else
        maxarr = myrow + counter
    End If

    'NOTE:  if mycol - 1 > mydata(counter), the entire string is returned
    'Check for the case where mycol - 1 < 0, ie = -1
    If (mycol = 0) Then
        mycol = 1 'To make mycol - 1 = 0
    End If

    While (counter < maxarr)
        TheData = left$(mydata(counter), mycol - 1) + Chr$(0)
        ptr = dce_element_add(ptr, TheData)
        counter = counter + 1
    Wend
    rv = dce_send_array(Socket, tag, ptr)

End Sub

Sub dce_push_char_Darr(Socket%, tag$, mydata() As String, myrow&, mycol&)
    Dim ptr&, counter&, maxarr&, TheData$, mysize&, rv%
    
    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If

    counter = LBound(mydata)
    mysize = UBound(mydata) - counter + 1
    If (myrow > mysize) Then
        TheData = "<OEC VB ERROR>  dce_push_char_Darr <" + tag + "> bad rowsize:  expected at least <" + Str$(myrow) + "> got <" + Str$(mysize) + ">"
        Call dce_log_str(TheData)
        Call dce_seterr(34, "dce_push_char_Darr")
        Exit Sub
    End If

    ptr = 0
    maxarr = myrow + counter

    'NOTE:  if mycol - 1 > mydata(counter), the entire string is returned
    'Check for the case where mycol - 1 < 0, ie = -1
    If (mycol = 0) Then
        mycol = 1 'To make mycol - 1 = 0
    End If

    While (counter < maxarr)
        TheData = left$(mydata(counter), mycol - 1) + Chr$(0)
        ptr = dce_element_add(ptr, TheData)
        counter = counter + 1
    Wend
    rv = dce_send_array(Socket, tag, ptr)

End Sub

Sub dce_push_char_Dstr(Socket%, tag$, mydata$, mylen&)
    Dim ANULL$, TheData$, rv%
    ANULL = Chr$(0)

    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If

    'Take care of the special case where mylen -1 = -1
    If (mylen = 0) Then
        mylen = 1
    End If

    TheData = left$(mydata, mylen - 1) + Chr$(0)
    If (InStr(TheData, ANULL) <> mylen) Then
        TheData = "<OEC VB ERROR>  dce_push_char_Dstr <" + tag + "> bad rowsize:  expected at least <" + Str$(mylen) + ">"
        Call dce_log_str(TheData)
        Call dce_seterr(34, "dce_push_char_Dstr")
        Exit Sub
    End If

    rv = dce_table_push(Socket, tag, TheData, mylen)
End Sub

Sub dce_push_char_str(Socket%, tag$, mydata$, mylen&)
    Dim ANULL$, TheData$, rv%
    ANULL = Chr$(0)

    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If

    'Take care of the special case where mylen -1 = -1
    If (mylen = 0) Then
        mylen = 1
    End If

    TheData = left$(mydata, mylen - 1) + Chr$(0)
    rv = dce_table_push(Socket, tag, TheData, InStr(TheData, ANULL))

End Sub

Sub dce_push_double(Socket%, tag$, mynum#)
    Dim mydata$, mylen&, rc%
    
    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If
    
    mydata = Str$(mynum) + Chr$(0)
    Call dce_vb2cdouble_str(mydata$)
    mylen = LenB(mydata) + 1
    rc = dce_table_push(Socket%, tag, mydata, mylen)
End Sub

Sub dce_push_double_Dstr(Socket%, tag$, mydata() As Double, myrow&)
    Dim TheData$, ptr&, counter&, maxarr&, mysize&, rv%
    
    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If
    
    ptr = 0
    counter = LBound(mydata)
    mysize = UBound(mydata) - counter + 1
    If (mysize < myrow) Then
        TheData = "<OEC VB ERROR>  dce_push_double_Dstr <" + tag + "> bad rowsize:  expected at least <" + Str$(myrow) + "> got <" + Str$(mysize) + ">"
        Call dce_log_str(TheData)
        Call dce_seterr(34, "dce_push_double_Dstr")
        Exit Sub
    Else
        maxarr = counter + myrow
    End If
    
    While (counter < maxarr)
        TheData = Str$(mydata(counter)) + Chr$(0)
        Call dce_vb2cdouble_str(TheData)
        ptr = dce_element_add(ptr, TheData)
        counter = counter + 1
    Wend
    rv = dce_send_array(Socket, tag, ptr)

End Sub

Sub dce_push_double_str(Socket%, tag$, mydata() As Double, myrow&)
    Dim TheData$, ptr&, counter&, maxarr&, mysize&, rv%
    
    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If
    
    ptr = 0
    counter = LBound(mydata)
    mysize = UBound(mydata) - counter + 1
    If (mysize < myrow) Then
        maxarr = counter + mysize
    Else
        maxarr = counter + myrow
    End If
    
    While (counter < maxarr)
        TheData = Str$(mydata(counter)) + Chr$(0)
        Call dce_vb2cdouble_str(TheData)
        ptr = dce_element_add(ptr, TheData)
        counter = counter + 1
    Wend
    rv = dce_send_array(Socket, tag, ptr)

End Sub

Sub dce_push_float(Socket%, tag$, mynum!)
    Dim mydata$, mylen&, rc%
    
    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If
    
    mydata = Str$(mynum) + Chr$(0)
    mylen = LenB(mydata) + 1
    rc = dce_table_push(Socket%, tag, mydata, mylen)
    
End Sub

Sub dce_push_float_Dstr(Socket%, tag$, mydata() As Single, myrow&)
    Dim TheData$, ptr&, counter&, maxarr&, mysize&, rv%
    
    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If
    
    ptr = 0
    counter = LBound(mydata)
    mysize = UBound(mydata) - counter + 1
    If (mysize < myrow) Then
        TheData = "<OEC VB ERROR>  dce_push_float_Dstr <" + tag + "> bad rowsize:  expected at least <" + Str$(myrow) + "> got <" + Str$(mysize) + ">"
        Call dce_log_str(TheData)
        Call dce_seterr(34, "dce_push_float_Dstr")
        Exit Sub
    Else
        maxarr = counter + myrow
    End If
    
    While (counter < maxarr)
        TheData = Str$(mydata(counter)) + Chr$(0)
        ptr = dce_element_add(ptr, TheData)
        counter = counter + 1
    Wend
    rv = dce_send_array(Socket, tag, ptr)
End Sub

Sub dce_push_float_str(Socket%, tag$, mydata() As Single, myrow&)
    Dim TheData$, ptr&, counter&, maxarr&, mysize&, rv%
    
    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If
    
    ptr = 0
    counter = LBound(mydata)
    mysize = UBound(mydata) - counter + 1
    If (mysize < myrow) Then
        maxarr = counter + mysize
    Else
        maxarr = counter + myrow
    End If
    
    While (counter < maxarr)
        TheData = Str$(mydata(counter)) + Chr$(0)
        ptr = dce_element_add(ptr, TheData)
        counter = counter + 1
    Wend
    rv = dce_send_array(Socket, tag, ptr)

End Sub

Sub dce_push_int_Dstr(Socket%, tag$, mydata() As Integer, myrow&)
    Dim TheData$, ptr&, counter&, maxarr&, mysize&, rv%
    
    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If
    
    ptr = 0
    counter = LBound(mydata)
    mysize = UBound(mydata) - counter + 1
    If (mysize < myrow) Then
        TheData = "<OEC VB ERROR>  dce_push_int_Dstr <" + tag + "> bad rowsize:  expected at least <" + Str$(myrow) + "> got <" + Str$(mysize) + ">"
        Call dce_log_str(TheData)
        Call dce_seterr(34, "dce_push_int_Dstr")
        Exit Sub
    Else
        maxarr = counter + myrow
    End If
    
    While (counter < maxarr)
        TheData = Str$(mydata(counter)) + Chr$(0)
        ptr = dce_element_add(ptr, TheData)
        counter = counter + 1
    Wend
    rv = dce_send_array(Socket, tag, ptr)
End Sub

Sub dce_push_int_str(Socket%, tag$, mydata() As Integer, myrow&)
    Dim TheData$, ptr&, counter&, maxarr&, mysize&, rv%
    
    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If
    
    ptr = 0
    counter = LBound(mydata)
    mysize = UBound(mydata) - counter + 1
    If (mysize < myrow) Then
        maxarr = counter + mysize
    Else
        maxarr = counter + myrow
    End If
    
    While (counter < maxarr)
        TheData = Str$(mydata(counter)) + Chr$(0)
        ptr = dce_element_add(ptr, TheData)
        counter = counter + 1
    Wend
    rv = dce_send_array(Socket, tag, ptr)

End Sub

Sub dce_push_long_Dstr(Socket%, tag$, mydata() As Long, myrow&)
    Dim TheData$, ptr&, counter&, maxarr&, mysize&, rv%
    
    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If
    
    ptr = 0
    counter = LBound(mydata)
    mysize = UBound(mydata) - counter + 1
    If (mysize < myrow) Then
        TheData = "<OEC VB ERROR>  dce_push_long_Dstr <" + tag + "> bad rowsize:  expected at least <" + Str$(myrow) + "> got <" + Str$(mysize) + ">"
        Call dce_log_str(TheData)
        Call dce_seterr(34, "dce_push_long_Dstr")
        Exit Sub
    Else
        maxarr = counter + myrow
    End If
    
    While (counter < maxarr)
        TheData = Str$(mydata(counter)) + Chr$(0)
        ptr = dce_element_add(ptr, TheData)
        counter = counter + 1
    Wend
    rv = dce_send_array(Socket, tag, ptr)
End Sub

Sub dce_push_long_str(Socket%, tag$, mydata() As Long, myrow&)
    Dim TheData$, ptr&, counter&, maxarr&, mysize&, rv%
    
    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If
    
    ptr = 0
    counter = LBound(mydata)
    mysize = UBound(mydata) - counter + 1
    If (mysize < myrow) Then
        maxarr = counter + mysize
    Else
        maxarr = counter + myrow
    End If
    
    While (counter < maxarr)
        TheData = Str$(mydata(counter)) + Chr$(0)
        ptr = dce_element_add(ptr, TheData)
        counter = counter + 1
    Wend
    rv = dce_send_array(Socket, tag, ptr)
End Sub

Sub dce_push_short_Dstr(Socket%, tag$, mydata() As Integer, myrow&)
    Dim TheData$, ptr&, counter&, maxarr&, mysize&, rv%
    
    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If
    
    ptr = 0
    counter = LBound(mydata)
    mysize = UBound(mydata) - counter + 1
    If (mysize < myrow) Then
        TheData = "<OEC VB ERROR>  dce_push_short_Dstr <" + tag + "> bad rowsize:  expected at least <" + Str$(myrow) + "> got <" + Str$(mysize) + ">"
        Call dce_log_str(TheData)
        Call dce_seterr(34, "dce_push_short_Dstr")
        Exit Sub
    Else
        maxarr = counter + myrow
    End If
    
    While (counter < maxarr)
        TheData = Str$(mydata(counter)) + Chr$(0)
        ptr = dce_element_add(ptr, TheData)
        counter = counter + 1
    Wend
    rv = dce_send_array(Socket, tag, ptr)
End Sub

Sub dce_push_short_str(Socket%, tag$, mydata() As Integer, myrow&)
    Dim TheData$, ptr&, counter&, maxarr&, mysize&, rv%
    
    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If
    
    ptr = 0
    counter = LBound(mydata)
    mysize = UBound(mydata) - counter + 1
    If (mysize < myrow) Then
        maxarr = counter + mysize
    Else
        maxarr = counter + myrow
    End If
    
    While (counter < maxarr)
        TheData = Str$(mydata(counter)) + Chr$(0)
        ptr = dce_element_add(ptr, TheData)
        counter = counter + 1
    Wend
    rv = dce_send_array(Socket, tag, ptr)
End Sub

Function dce_push_string%(Socket%, tag$, mydata$)
    Dim ANULL$, TheData$, mylen&, rv%
    
    ANULL = Chr$(0)

    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Function
    End If
    
    mylen = InStr(mydata, ANULL)
    If mylen = 0 Then
        TheData = mydata + ANULL
        mylen = LenB(mydata) + 1
    Else
        TheData = left$(mydata, mylen)
    End If
    
    rv = dce_table_push(Socket, tag, TheData, mylen)

End Function

Sub dce_push_void_Dstr(Socket%, tag$, mydata$, mylen&)
    Dim voidlen&, TheData$, rv%

    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If

    voidlen = LenB(mydata)
    If (voidlen < mylen) Then
        TheData = "<OEC VB ERROR>  dce_push_void_Dstr <" + tag + "> bad rowsize:  expected at least <" + Str$(mylen) + ">"
        Call dce_log_str(TheData)
        Call dce_seterr(34, "dce_push_void_Dstr")
        Exit Sub
    End If

    rv = dce_table_push(Socket, tag, mydata, mylen)

End Sub

Sub dce_push_void_str(Socket%, tag$, mydata$, mylen&)
    Dim voidlen&, TheData$, rv%

    'First check the library error flag
    If (dce_iserror() <> 0) Then
        Exit Sub
    End If

    'For voids, you must push the exact amount
    voidlen = LenB(mydata)
    If (voidlen < mylen) Then
        TheData = "<OEC VB ERROR>  dce_push_void_str <" + tag + "> bad rowsize:  expected at least <" + Str$(mylen) + ">"
        Call dce_log_str(TheData)
        Call dce_seterr(34, "dce_push_void_str")
        Exit Sub
    End If

    rv = dce_table_push(Socket, tag, mydata, mylen)
End Sub

Sub dce_showerror()
    Dim ErrMsg As String * 250
    Dim rc%

    rc = dce_error(ErrMsg)
    rc = MsgBox(ErrMsg, 48, "DCE Error")
End Sub

Sub dce_start()
    Dim rv%
    rv = dce_setenv("c:\basicvb\client.env", "", "")
    If (rv = 0) Then
        Call dce_showerror
    End If
End Sub

Sub dce_vb2cdouble_str(double_str$)
'This function converts ###D### to ###e###

Dim exp_marker_pos%
exp_marker_pos = InStr(double_str, "D")
If exp_marker_pos > 0 Then
    double_str = Mid$(double_str, 1, exp_marker_pos - 1) + "e" + Mid$(double_str, exp_marker_pos + 1)
End If

End Sub

Sub list2arr(mylist As Control, myarray() As String)
    Dim i%
    ReDim myarray(0 To mylist.ListCount - 1)
    For i = 0 To mylist.ListCount - 1
        myarray(i) = mylist.List(i)
    Next i
End Sub

