<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<%
f_permission = administrator(false,session("user"),"m_customer")
if f_permission < 2 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
	status	=	Request.QueryString("status")
	action	=	GetNumeric(Request.Form("action"),0)
	ID = 0
if action = 0  then
	if status = "del" then
		ID	=	GetNumeric(Request.QueryString("ID"),0)
		sql = "Delete TaiKhoan where id="&ID
		set rsTK 	=	Server.CreateObject("ADODB.recordset")	
		rsTK.open sql,con,1
		set rsTK = nothing
		%>		
		<script language="javascript">
			window.close();
			window.opener.reload();
		</script>
		<%		
		Response.End()
	elseif status = "edit" then
		ID	=	GetNumeric(Request.QueryString("ID"),0)
		sql = "SELECT * from TaiKhoan where id='"&ID&"'"
		set rsTK 	=	Server.CreateObject("ADODB.recordset")
		rsTK.open sql,con,1
		if not rsTK.eof then
			Ngay	=	GetNumeric(Day(rsTK("iniDates")),1)
			Thang	=	GetNumeric(Month(rsTK("iniDates")),1)
			Nam		=	GetNumeric(Year(rsTK("iniDates")),1900)
			Lydo	=	rsTK("Lydo")
			iniTK	=	rsTK("iniTK")
		end if
		set rsTK = nothing
		TitleTK	=	"Sửa tài khoản"
	else
		Ngay	=	Day(now)	
		Thang	=	Month(now)
		Nam		=	Year(now)
		Lydo	=	""
		iniTK	=	0	
		TitleTK	=	"Thêm tài khoản"	
	end if
else
	Ngay	=	GetNumeric(Request.Form("Ngay"),1)
	Thang	=	GetNumeric(Request.Form("Thang"),1)
	Nam		=	GetNumeric(Request.Form("Nam"),1900)
	Lydo	=	Trim(Request.Form("txtLydo"))
	iniTK	=	Chuan_money(Request.Form("txtTien"))
	dates	=	Thang&"/"&Ngay&"/"&Nam
	if status = "edit" then
		ID	=	GetNumeric(Request.Form("ID"),0)
		sql	=	"update TaiKhoan Set iniTK = '"& iniTK &"', iniDates='"&dates&"',Lydo=N'"&Lydo&"' where ID="&ID	
	else
		CMND	=	Request.Form("CMND")
		sql	=	"insert into TaiKhoan(CMND,iniTK,iniDates,Lydo) values('"& CMND &"','"& iniTK &"','"& dates &"',N'"& Lydo &"')"
	end if
	set rsTK = Server.CreateObject("ADODB.recordset")
	
	rsTK.open sql,con,1
	%>		
	<script language="javascript">
		window.close();
		window.opener.history.back();
		window.opener.reload();
	</script>
<%	
Response.End()
end if	
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Edit Customer</title>
<link href="../../css/styles.css" rel="stylesheet" type="text/css">
</head>
<script language="JavaScript" src="include/vietuni.js"></script>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form action="EditTKCustomer.asp?status=<%=status%>" target="_blank" method="post">
<table width="99%" border="0" cellpadding="0" cellspacing="0" class="CTxtContent">
  <tr>
    <td colspan="2" height="30" background="../../images/TabChinh.gif" style="background-repeat:no-repeat" valign="bottom">&nbsp;&nbsp;&nbsp;&nbsp;<span class="CTieuDeNho"><%=TitleTK%></span></td>
  </tr>
  <tr>
    <td width="16%">Ngày:</td>
    <td width="84%">
	                  <%
					Call List_Date_WithName(Ngay,"DD","Ngay")
					Call List_Month_WithName(Thang,"MM","Thang")
					Call List_Year_WithName(Nam,"YYYY",2004,"Nam")
				%>	</td>
  </tr>
  <tr>
    <td>Lý do: </td>
    <td><textarea name="txtLydo" cols="30" rows="3" id="txtLydo"><%=Lydo%></textarea></td>
  </tr>
  <tr>
    <td>Số tiền: </td>
    <td><input name="txtTien" type="text" id="txtTien" onKeyUp="javascript: DisMoneyThis(this);" size="20" value="<%=Dis_str_money(iniTK)%>">
      đ</td>
  </tr>
  <tr>
    <td colspan="2" align="center">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" align="center">
	<input name="action" value="1" type="hidden">
	<input name="ID" value="<%=ID%>" type="hidden">
	<input name="CMND" value="<%=Request.QueryString("CMND")%>" type="hidden">
	
	<input name="Submit" type="submit" id="Submit" value="    OK    ">

      <input type="reset" name="Submit2" value="   Reset   ">
      <input type="reset" name="Submit22" value=" Quay lại " onClick="javascript:history.back();"></td>
  </tr>
</table>
</form>
</body>

</html>
