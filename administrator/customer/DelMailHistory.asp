<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<%
    
    ID	=	GetNumeric(Request.QueryString("ID"),0)
    Delall = Trim(Request.QueryString("Delall"))
    if ID <> 0 then
        sql = "delete EmailCompose where ID="&ID	
    elseif Delall = "all" then
        sql = "delete EmailCompose"
    end if
    
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.open sql,Con,3
    set rs = nothing

     %>
	<script language="javascript">
	    history.back();
	    window.opener.location.reload();
	</script>