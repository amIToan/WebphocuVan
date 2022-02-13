<%session.CodePage=65001%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<%
f_permission = administrator(false,session("user"),"m_user")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<html>
	<head>
	<title><%=PAGE_TITLE%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
    <link href="../../css/styles.css" rel="stylesheet" type="text/css">
    <LINK href="/administrator/inc/admstyle.css" type=text/css rel=stylesheet>    
    <SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
    <LINK href="/administrator/inc/admstyle.css" type=text/css rel=stylesheet>
    <script type="text/javascript" src="/administrator/inc/common.js"></script>
    <link href="/administrator/inc/admstyle.css" type="text/css" rel="stylesheet">
    <link href="/administrator/skin/css/sweetalert.css" rel="stylesheet" />
    <link href="/administrator/skin/css/bootstrap.css" rel="stylesheet" />
    <script src="/administrator/skin/script/jquery-2.2.2.min.js"></script>
    <script src="/administrator/skin/script/ajax-asp.js"></script>
    <link href="/administrator/css/skin1.css" rel="stylesheet" />
    <link type="text/css" href="../../css/w3.css" rel="stylesheet" />
	</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div class="container-fluid">
<%
	Title_This_Page="Quản lý -> Người sử dụng"
	Call header()
	
%>
<div class="container-fluid">
    <div class="col-md-2" style="background:#001e33;"><%call MenuVertical(0) %></div>
    <div class="col-md-10">
<table width="770" border="0" align="center" cellpadding="0" cellspacing="0">
<%if f_permission >=1 then%>
  <tr align="right" valign="top"> 
    <<!--td height="25"> 
	<font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>
	<a href="javascript: winpopup('user_update.asp','',900,600);">Tạo User mới</a></strong>
	</font></td>-->
  </tr>
 <%end if%>
<!--  <tr align="center"> 
    <td height="25" valign="middle"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
		<%c=Request.QueryString("c")
			if Trim(c)="" or not isnumeric(c) then
				c=65
			else
				c=Clng(c)
			end if
			for i=65 to 90
				if i<>c then
					response.Write("<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?c=" & i & """ style=""text-decoration:none"">"& chr(i) & "</a>|")
				else
					response.Write("<font color=""red""><strong>" & chr(i) & "</strong></font>|")
				end if
			Next
			
			if c<>0 then
				response.Write("<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?c=0"" style=""text-decoration:none"">T&#7845;t c&#7843;</a>")
			else
				response.Write("<font color=""red""><strong>T&#7845;t c&#7843;</strong></font>")
			end if
		%></font></td>
  </tr>-->
</table>
<%
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	sql="SELECT [User].UserName, [User].CreationDate, [User].LastLoginDate, Nhanvien.Ho_Ten FROM [User] INNER JOIN Nhanvien ON [User].IDNhanVien = Nhanvien.NhanVienID"
	sql=sql & " order by LastLoginDate desc"
	rs.PageSize = 50
	rs.open sql,con,1
	
	if rs.eof then 'Không có bản ghi nào thỏa mãn
		Response.Write "<table width=""770"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewline &_
							"<tr align=""left"">" & vbNewline &_
		                       "<td height=""60"" valign=""middle""><strong><font size=""2"" face=""Verdana, Arial, Helvetica, sans-serif"">Kh&#244;ng c&#243; d&#7919; li&#7879;u</font></strong></td>" & vbNewline &_
							"</tr>"& vbNewline &_
						"</table>" & vbNewline
	else
		if request.Querystring("page")<>"" then
			page=Clng(request.Querystring("page"))
		else
			page=1
		end if

		rs.AbsolutePage = CLng(page)
		i=0
		stt=(page-1)* rs.pageSize + 1
%>
<table width="770" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#000000" class="CTxtContent" >
  <tr align="center" bgcolor="FFFFFF"> 
  	<td><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">TT</font></strong></td>
    <td><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Tên 
      truy nhập</font></strong></td>
    <td><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Họ tên</font></strong></td>
    <td><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Ngày 
      tạo</font></strong></td>
    <td><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Đăng nhập<br>
      gần nhất</font></strong></td>
    <td>&nbsp;</td>
  </tr>
  <%Do while not rs.eof and i<rs.pagesize%>
  <tr 	<%if i mod 2=0 then response.Write("bgcolor=""#E6E8E9""") else response.Write("bgcolor=""#FFFFFF""")%>>
  	<td align="right" valign="middle"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=stt%>.&nbsp;</font></td>
    <td valign="middle"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("UserName")%></font>
	<%if CheckUserRoleExist(rs("UserName"))=0 then%>
		<br><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FF0000">* Inactive</font>
	<%End if%></td>
    <td valign="middle"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("Ho_Ten")%></font></td>
    <td align="center" valign="middle"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
		<%=Day(ConvertTime(rs("CreationDate")))%>/<%=Month(ConvertTime(rs("CreationDate")))%>/<%=Year(ConvertTime(rs("CreationDate")))%>
	</font></td>
    <td align="center" valign="middle"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
		<%if not IsNull(rs("LastLoginDate")) then%>
			<%=Hour(ConvertTime(rs("LastLoginDate")))%>:<%=Minute(ConvertTime(rs("LastLoginDate")))%>' <%=Day(ConvertTime(rs("LastLoginDate")))%>/<%=Month(ConvertTime(rs("LastLoginDate")))%>/<%=Year(ConvertTime(rs("LastLoginDate")))%>
		<%Else%>
			&nbsp;
		<%End if%>
	</font></td>
    <td align="center" valign="middle"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">
		<%if f_permission > 1 then%>
		<a href="javascript: winpopup('user_update.asp','<%=rs("Username")%>&act=edit',900,600);">Sửa</a>
		<%end if%>
		<%if f_permission > 2 then%>
		<a href="javascript: winpopup('user_delete.asp','<%=rs("Username")%>',300,150);">|Xóa</a>
		<%end if%>
	</font></td>
  </tr>
  <%i=i+1
  stt=stt+1
  rs.movenext
  Loop
  pagecount=rs.pagecount
  pageperbook=7
  Response.Write"<tr><td colspan=""7"" align=""center"" bgcolor=""#FFFFFF"">"
  		Call phantrang(page,pagecount,pageperbook)
  Response.Write "</td></tr>"
End if
rs.close
set rs=nothing%>
</table>
</div>
    </div>
        </div>
<%Call Footer()%>
</body>
</html>
