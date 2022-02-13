<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
f_permission = administrator(false,session("user"),"m_sys")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
	action =trim(Request.Querystring("action"))
	action=replace(action,"'","''")
	if action = "" or action=Null then
		action =trim(Request.Form("action"))
		action=replace(action,"'","''")
	end if
%>
<html>
<head>

<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="../../css/styles.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<p>
  <%
select case action
  case "update"
  	set rs=server.CreateObject("ADODB.Recordset")
	sql="SELECT Count(id) as iCount FROM XSEOPrice "
	rs.open sql,con,1
	i =	rs("iCount")
	j=1
	for j= 1 to i
		id      =   Trim(Request.Form("ID"&j))
		Title   =   Trim(Request.Form("Title"&j))
		Download=   Trim(Request.Form("Download"&j))

        Note    =   Trim(Request.Form("Note"&j))
        Note    =   Replace(Note,"'","''")
        
		Price   =   Chuan_money(Request.Form("Price"&j))

        PriceOff   =   Chuan_money(Request.Form("PriceOff"&j))

		fMonth  =   Trim(Request.Form("Month"&j))

        Description=Trim(Request.Form("Description"&j))
		Description=Replace(Description,"'","''")

		sql 	=	"update XSEOPrice set Title=N'"& Title &"', Download ='"& Download &"',Note = N'"& Note &"', Price = '"& Price &"', PriceOff='"& PriceOff &"',Month = '"& fMonth &"', Description=N'"& Description &"'"
		sql     =   sql + " where id = '"&id&"'"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.open sql,con,1
		set	rs = nothing
       
      
	next

	i=i+1

	Title   =   Trim(Request.Form("Title"&i))
    if Title <> "" then
	    Download=   Trim(Request.Form("Download"&i))

        Note    =   Trim(Request.Form("Note"&i))
        Note    =   Replace(Note,"'","''")
       
	    Price   =   Chuan_money(Request.Form("Price"&i))
	    fMonth  =   Trim(Request.Form("Month"&i))

        Description=Trim(Request.Form("Description"&i))
	    Description=Replace(Description,"'","''")


	    sql =	"insert into XSEOPrice(Title, Download, Note, Price, PriceOff,Month,Description) values(N'"& Title &"','"& Download &"',N'"& Note &"','"& Price &"','"& PriceOff &"','"& fMonth &"',N'"& Description &"')"
	    Set rs = Server.CreateObject("ADODB.Recordset")
	    rs.open sql,con,1
	    set	rs = nothing
    end if
	Response.Write	"<script language=""JavaScript"">" & vbNewline &_
			"	<!--" & vbNewline &_
			"		window.opener.location.reload();" & vbNewline &_
			"		window.close();" & vbNewline &_	
			"	//-->" & vbNewline &_
			"</script>" & vbNewline
	%>
	<script language="javascript">
	    history.back();
	</script>
	<%
 case "del"
 	id=Request.QueryString("id")
	sql="delete XSEOPrice where id='" & id & "'"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
	set	rs = nothing
	Response.Write	"<script language=""JavaScript"">" & vbNewline &_
			"	<!--" & vbNewline &_
			"		window.opener.location.reload();" & vbNewline &_
			"		window.close();" & vbNewline &_
			"	//-->" & vbNewline &_
			"</script>" & vbNewline
end select

%>
</p>
</body>
</html>
