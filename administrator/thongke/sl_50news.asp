<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<%
f_permission = administrator(false,session("user"),"m_report")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%IF Request.form("action")="Search" then
	Ngay1=GetNumeric(Request.form("Ngay1"),0)
	Thang1=GetNumeric(Request.form("Thang1"),0)
	Nam1=GetNumeric(Request.form("Nam1"),0)
	Ngay2=GetNumeric(Request.form("Ngay2"),0)
	Thang2=GetNumeric(Request.form("Thang2"),0)
	Nam2=GetNumeric(Request.form("Nam2"),0)

ELSE
	Ngay1=Day(now())
	Thang1=Month(now())-1
	Nam1=Year(now())
	Ngay2=Day(now())
	Thang2=Month(now())
	Nam2=Year(now())
END IF

FromDate=Thang1 & "/" & Ngay1 & "/" & Nam1
ToDate=Thang2 & "/" & Ngay2 & "/" & Nam2
%>
<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
<LINK href="/administrator/inc/admstyle.css" type=text/css rel=stylesheet>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
	Call header()
	Call Menu()
	Title_This_Page="Th&#7889;ng k&#234; -> 50 tin được đọc nhiều nhất"
	
%>
<FORM action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="post" name="fThongke" onSubmit="return checkme();">
  <table align="center" cellpadding="0" cellspacing="0" width="770">
    <tr> 
      <td align="right" valign="middle"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Chọn khoảng thời gian:</strong></font>
		<%
			Call List_Date_WithName(Ngay1,"DD","Ngay1")
			Call List_Month_WithName(Thang1,"MM","Thang1")
			Call  List_Year_WithName(Nam1,"YYYY",2004,"Nam1")
		%>
        <img src="../images/right.jpg" width="9" height="9" align="absmiddle"> 
        <%
			Call List_Date_WithName(Ngay2,"DD","Ngay2")
			Call List_Month_WithName(Thang2,"MM","Thang2")
			Call  List_Year_WithName(Nam2,"YYYY",2004,"Nam2")
		%>
		<input type="image" name="ButtonSearch" id="ButtonSearch" align="absmiddle" src="/administrator/images/search_bt.gif" width="23" height="23" border="0">
		<input type="hidden" name="action" value="Search">
		<input type="hidden" name="OrderType" value="">
	  </td>
    </tr>
  </table>
</form>
<SCRIPT LANGUAGE=JavaScript>
<!--
 function order(OrderType)
 {
 	if (!checkme())
 		return;
 	document.fThongke.OrderType.value=OrderType;
 	document.fThongke.submit();
 }
 function checkme()
 {
 	if (document.fThongke.Ngay1.value==0)
	{
		alert("Bạn chưa chọn ngày!");
		document.fThongke.Ngay1.focus();
		return false;
	}
	if (document.fThongke.Thang1.value==0)
	{
		alert("Bạn chưa chọn tháng!");
		document.fThongke.Thang1.focus();
		return false;
	}
	if (document.fThongke.Nam1.value==0)
	{
		alert("Bạn chưa chọn năm!");
		document.fThongke.Nam1.focus();
		return false;
	}
	if (document.fThongke.Ngay2.value==0)
	{
		alert("Bạn chưa chọn ngày!");
		document.fThongke.Ngay2.focus();
		return false;
	}
	if (document.fThongke.Thang2.value==0)
	{
		alert("Bạn chưa chọn tháng!");
		document.fThongke.Thang2.focus();
		return false;
	}
	if (document.fThongke.Nam2.value==0)
	{
		alert("Bạn chưa chọn năm!");
		document.fThongke.Nam2.focus();
		return false;
	}
	return true;
 }
// -->
</SCRIPT>
<%
IF Request.form("action")="Search" And IsDate(ToDate) AND IsDate(FromDate) THEN
	FromDate=FormatDateTime(FromDate)
	ToDate=FormatDateTime(ToDate)
  	Dim rs
  	Set rs=Server.CreateObject("ADODB.Recordset")
  	
	sql="SELECT top 50 n.NewsId, n.Title,n.Creator,n.CreationDate, n.NewsCount"
		sql=sql & " FROM News n"
		sql=sql & " WHERE (DATEDIFF(dd, n.CreationDate, '" & FromDate & "') <= 0) AND (DATEDIFF(dd, n.CreationDate, '" & ToDate & "') >= 0) "
		sql=sql & " ORDER BY n.NewsCount desc"
	rs.open sql,con,3
%>
<table width="770" align="center" cellpadding="0" cellspacing="1" bordercolor="#000000" bgcolor="#000000">
  <tr bgcolor="#FFFFFF"> 
    <td><div align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">TT</font></strong></div></td>
    <td><div align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Tiêu đề tin</font></strong></div></td>
	<td><div align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Lượt<br>xem</font></strong></div></td>
  </tr>
<%stt=1
  Do while not rs.eof
%>
  <tr bgcolor="#FFFFFF"> 
    <td align="right" valign="top"><font size="2" face="Arial, Helvetica, sans-serif"><%=stt%>.</font></td>
    <td><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;<a href="javascript: winpopup('/administrator/news/news_view.asp','<%=rs("NewsId")%>&CatId=1',600,400);" style="text-decoration: none"><%=rs("Title")%></a><br>
      <font size="1">&nbsp;(Tạo bởi: <%=rs("Creator")%>-<%=Hour(ConvertTime(rs("CreationDate")))%>h<%=Minute(ConvertTime(rs("CreationDate")))%>&quot;&nbsp;<%=Day(ConvertTime(rs("CreationDate")))%>/<%=Month(ConvertTime(rs("CreationDate")))%>/<%=Year(ConvertTime(rs("CreationDate")))%>)</font>
    </font></td>
	<td align="center" valign="middle"><font size="2" face="Arial, Helvetica, sans-serif"><%=rs("NewsCount")%></font></td>
  </tr>
<%
stt=stt+1
rs.movenext
Loop
rs.close
set rs=nothing%>
</table>

<%
END IF 'IF Request.form("action")="Search" THEN
%>
<%Call Footer()%>
</body>
</html>