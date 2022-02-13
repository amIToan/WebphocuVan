<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/Fs_liblary.asp" -->
<%
    key_  =  Request.Form("_key")
    id_  =  Request.Form("_id")
    stt_status = 0
    IF key_ <> "" And  key_ = "Del" THEN  
    sql  = "DELETE Company WHERE ID =  '"&id_&"'"
        on error resume next
        con.Execute sql,recaffected      
        if err<>0 then 'not ok
           stt_status = 0
        else 'ok
           stt_status = 1
        end if
        conn.close
        Response.Write stt_status
    END IF
%>