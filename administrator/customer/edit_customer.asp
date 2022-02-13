<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<%
f_permission = administrator(false,session("user"),"m_customer")
if f_permission < 2 then
	response.Redirect("/administrator/info.asp")
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Edit Customer</title>
</head>
<script language="JavaScript" src="include/vietuni.js"></script>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
	cmnd=Request.QueryString("cmnd")
	Dim i 
  	msg = "Bạn có chắc chắn sửa dữ liệu?"
   	Response.Write("<" & "script language=VBScript>")
    i=	Response.Write("MsgBox """ & msg & """,4, ""Xin hỏi""</script>")
	i=6
	if i = 6 then 
		f_cmnd_gt = Trim(Request.Form("f_cmnd_gt"))
		f_pass    = Trim(Request.Form("f_pass"))
		f_name    = Trim(Request.Form("f_name"))
		f_quequan = Trim(Request.Form("f_quequan"))
		f_diachi  = Trim(Request.Form("f_diachi"))
		f_mail    = Trim(Request.Form("f_mail"))
		f_tell    = Trim(Request.Form("f_tell"))
		f_mobile  = Trim(Request.Form("f_mobile"))
		Dim rs
		sql="update Account set nguoi_gioi_thieu = N'"& f_cmnd_gt &"',"
		sql=sql    +		"password         =  '"& f_pass &"',"
		sql=sql    +		"Name             = N'"& f_name &"',"
		sql=sql    +		"nguyenquan       = N'"& f_quequan &"',"
		sql=sql    +		"diachi           = N'"& f_diachi &"',"
		sql=sql    +		"Email            = N'"& f_mail &"',"
		sql=sql    +		"Tell             = N'"& f_tell &"',"
		sql=sql    +		"mobile           = N'"& f_mobile &"'"
		sql=sql    +        "Where CMND="&cmnd
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.open sql,con,1
		set rs=nothing
	end if
		response.Write "<script language=""JavaScript"">" & vbNewline &_
				"<!--" & vbNewline &_
				"window.opener.location.reload();" & vbNewline &_
				"window.close();"&vbNewline&_
			"//-->" & vbNewline &_
			"</script>" & vbNewline
%>


</body>
</html>
