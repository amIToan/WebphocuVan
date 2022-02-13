<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/include/func_tiny.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<%
f_permission = administrator(false,session("user"),"m_report")
if f_permission < 2 then
	response.Redirect("/administrator/info.asp")
end if
%>

<%
actions	=	Request.form("action")
IF actions="Search" then
	Ngay1=GetNumeric(Request.form("Ngay1"),0)
	Thang1=GetNumeric(Request.form("Thang1"),0)
	Nam1=GetNumeric(Request.form("Nam1"),0)
	Ngay2=GetNumeric(Request.form("Ngay2"),0)
	Thang2=GetNumeric(Request.form("Thang2"),0)
	Nam2=GetNumeric(Request.form("Nam2"),0)
	strKhoHang	=	Trim(Request.Form("txtKhoHang"))
	strSelSearch	=	Trim(Request.Form("selSearch"))
	iOrderBy	=  Clng(Request.Form("RaOderBy"))
	

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
<link href="../../css/styles.css" rel="stylesheet" type="text/css"></head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%	Title_This_Page="Thống kê -> kho hàng"
	Call header()
	Call Menu()

	
%>
<FORM action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="post" name="fThongke" onSubmit="return checkme();">
  <table width="100%" align="center" cellpadding="1" cellspacing="1" bgcolor="#FFFF99"  style="border:#666666 solid 1">
    <tr>
      <td width="209" align="right" valign="middle"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Tại kho hàng:</strong></font></td>
      <td width="221" align="right" valign="middle"><div align="left">
        <input name="txtKhohang" type="text" class="CTextBoxUnder" size="20" value="<%=strKhoHang%>">     
      </div></td> 
      <td width="563" align="right" valign="middle"> <div align="left"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Thời gian:</strong></font>
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
      </div></td>
    </tr>
    <tr>
      <td align="right" valign="middle"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Sắp xếp theo:</strong></font></td>
      <td align="right" valign="middle"><div align="left"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>
        <select name="selSearch">
          <option value="CreationDate" <%if strSelSearch = "CreationDate" then Response.Write("selected") end if %> >Theo ngày</option>
          <option value="Title" <%if strSelSearch = "Title" then Response.Write("selected") end if %>>Theo tiêu đề tin</option>
          <option value="Giabia" <%if strSelSearch = "Giabia" then Response.Write("selected") end if %>>Theo giá</option>
          <option value="nxb" <%if strSelSearch = "nxb" then Response.Write("selected") end if %>>Theo nxb</option>
          <option value="Creator" <%if strSelSearch = "Creator" then Response.Write("selected") end if %>>Biên tập viên</option>
        </select>
      </strong></font></div></td>
      <td align="left" valign="middle"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> Sắp xếp: </strong></font><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Tăng dần</strong></font>
        <input name="RaOderBy" type="radio" value="0" <%if iOrderBy =0 then Response.Write("checked") end if%>>
        <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>/Giảm dần
        <input name="RaOderBy" type="radio" value="1" <%if iOrderBy =1 then Response.Write("checked") end if%>>
      </strong></font></td>
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
<%if actions="Search" then
		sqlEx=sqlEx & " FROM News n INNER JOIN NewsDistribution d ON n.NewsID = d.NewsID INNER JOIN NewsCategory c ON d.CategoryID = c.CategoryID"
		sqlEx=sqlEx & " WHERE  n.Tinhtrang = N'" & strKhoHang & "' and c.CategoryLoai = 3"
		sqlEx=sqlEx & " and (DATEDIFF(dd, n.CreationDate, '" & FromDate & "') <= 0) AND (DATEDIFF(dd, n.CreationDate, '" & ToDate & "') >= 0) "		
		sql1	=	"SELECT Count(n.NewsID) as numTotal, SUM(n.Gia) as iTotalCK,SUM(n.Giabia) as iTotalBia  "
		sql1	=	sql1 + sqlEx
		set rsTotal=server.CreateObject("ADODB.Recordset")
		rsTotal.open sql1,con,1
		if not rsTotal.eof and rsTotal("numTotal")<>0 then
			
%>
<table width="585" border="0" cellpadding="0" cellspacing="1" bgcolor="#FFFFCC" class="CTxtContent" style="border:#999999 solid 1">
  <tr> 
    <td width="185" align="right"><u>Danh sách</u></td>
    <td width="395" align="left">từ ngày <b><%=Ngay1%>/<%=Thang1%>/<%=Nam1%></b> đến ngày <b><%=Ngay2%>/<%=Thang2%>/<%=Nam2%></b>	</td>
  </tr>
 
  <tr> 
    <td align="right"><u>Tổng cộng có:</u></td>
    <td align="left"><b><%=rsTotal("numTotal")%></b> cuốn</font></td>
  </tr>
  <tr>
    <td align="right"><u>Tổng giá theo bìa: </u></td>
    <td align="left"><%=Dis_str_money(rsTotal("iTotalBia"))&DonviGia %></td>
  </tr>
  <tr>
    <td align="right"><u>Tổng giá sau chiết khẩu: </u></td>
    <td align="left"><%=Dis_str_money(rsTotal("iTotalCK"))&DonviGia%></td>
  </tr>
</table>
<br>
<%
		end if
		if iOrderBy = 1 then 
			sqlEx=sqlEx & " ORDER BY "& strSelSearch &" desc"
		else
			sqlEx=sqlEx & " ORDER BY "& strSelSearch 
		end if

		sql="SELECT n.*, c.* " + sqlEx
		
	set rs=server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
	iSTT=1

  %>
  
  <table width="100%" align="center" cellpadding="0" cellspacing="1" bordercolor="#000000" bgcolor="#000000" class="CTxtContent">
  <tr bgcolor="#FFFFFF">
    <td width="23" align="center"><strong>TT</strong></td>
    <td width="311" align="center"><strong>Tiêu đề</strong></td>
    <td width="201"align="center"><strong>Tác giả </strong></td>
    <td width="146" align="center"><strong>NXB </strong></td>
    <td width="91" align="center"><strong>User</strong></td>
    <td width="91" align="center"><strong> Giá bìa </strong></td>
    <td width="29" align="center"><strong>SL</strong></td>
    <td width="92" align="center"><strong>Sau CK% </strong></td>
  </tr>
  <%
  Do while not rs.eof
	giabia	=	rs("Giabia")
	gia		=	rs("Gia")
%>
  <tr bgcolor="#FFFFFF">
    <td align="right" valign="top"><%=iSTT%>.</td>
    <td>
		<%
		 	Title = rs("Title")
			i = len(Title)
			temp=""
			if i > 38 then
				temp	=	Left(Title,38)
				temp	=	temp+"..."
			else
				temp =Title
			end if
		 %>  
		 <a title="<%=Title%>"> <%=temp%> </a>
	</td>
    <td align="left" valign="middle"> 
		<%
		 	tacgia = rs("tacgia")
			i = len(tacgia)
			temp=""
			if i > 28 then
				temp	=	Left(tacgia,28)
				temp	=	temp+"..."
			else
				temp = tacgia				
			end if
		 %>  
		 <a title="<%=tacgia%>"> <%=temp%> </a>
	</td>
    <td align="left" valign="middle">
		<%
		 	nxb = rs("nxb")
			i = len(nxb)
			temp=""
			if i > 20 then
				temp	=	Left(nxb,20)
				temp	=	temp+"..."
			else
				temp = nxb
			end if
		 %>  
		 <a title="<%=nxb%>"> <%=temp%> </a>
	</td>
    <td align="center" valign="middle"><%=rs("Creator")%></td>
    <td align="center" valign="middle"><div align="right"><%=giabia%></div></td>
    <td align="center" valign="middle">1</td>
    <td align="center" valign="middle"><div align="right"><%=gia%></div></td>
  </tr>
<%	
	iSTT=iSTT+1
rs.movenext
Loop
%>
  <form action="<%=Request.ServerVariables("Script_name")%>" method="post" name="fSearch2">
    <input type="hidden" name="ngay1" value="<%=ngay1%>">
    <input type="hidden" name="thang1" value="<%=thang1%>">
    <input type="hidden" name="nam1" value="<%=nam1%>">
    <input type="hidden" name="ngay2" value="<%=ngay2%>">
    <input type="hidden" name="thang2" value="<%=thang2%>">
    <input type="hidden" name="nam2" value="<%=nam2%>">
    <input type="hidden" name="ViewType" value="<%=ViewType%>">
    <input type="hidden" name="Username" value="<%=Username%>">
    <input type="hidden" name="page">
    <input type="hidden" name="action2" value="Search">
  </form>
	
</table>

<br>

<%	
	rs.close
	set rs=nothing
end if' search
%>

<%Call Footer()%>
</body>
</html>