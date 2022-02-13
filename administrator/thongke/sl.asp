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
	
	OrderType=ReplaceHTMLToText(Request.form("OrderType"))
ELSE
	Ngay1=Day(now())
	Thang1=Month(now())-1
	Nam1=Year(now())
	Ngay2=Day(now())
	Thang2=Month(now())
	Nam2=Year(now())
END IF

%>
<html>
<head>
    <title><%=PAGE_TITLE%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
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
	Title_This_Page="Thống kê -> SL tin theo Biên tập viên"
	Call header()	
%>
<div class="container-fluid">
    <div class="col-md-2" style="background:#001e33;"><%call MenuVertical(0) %></div>
    <div class="col-md-10">
<FORM action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="post" name="fThongke" onSubmit="return checkme();">
  <table align="center" cellpadding="0" cellspacing="0" width="770" class="w3-table w3-table-all w3-round w3-margin">
    <tr> 
      <td align="right" valign="middle"> <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Thời gian:</strong></font>
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
		<input type="hidden" name="ButtonSearch" id="ButtonSearch" align="absmiddle" src="/administrator/images/search_bt.gif" width="23" height="23" border="0">
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
FromDate=Thang1 & "/" & Ngay1 & "/" & Nam1
ToDate=Thang2 & "/" & Ngay2 & "/" & Nam2

'IF Request.form("action")="Search" and IsDate(FromDate) and IsDate(ToDate) THEN
  	FromDate=FormatDatetime(FromDate)
	ToDate=FormatDatetime(ToDate)
	
	sqlUser="SELECT	u.UserName " &_
			"FROM	[User] u INNER JOIN " &_
            "		UserDistribution ud ON u.UserName = ud.UserName " &_
			"WHERE	(ud.User_role = 'ap') OR (ud.User_role = 'ad') OR (ud.User_role = 'ed') OR (ud.User_role = 'se') " &_
			"GROUP BY u.UserName "
	if  session("LstRole")="0ad" then
	else
		sqlUser=sqlUser & "HAVING (u.UserName = '" & session("user") & "') " 
	end if
	sqlUser=sqlUser & "ORDER BY u.UserName"
	
	Dim rsUser
	Set rsUser=Server.CreateObject("ADODB.recordset")
	rsUser.open sqlUser,con,3
	
	
%>
<form name="fView" method="post" action="sl_view.asp">
	<input type="hidden" name="ngay1" value="">
	<input type="hidden" name="thang1" value="">
	<input type="hidden" name="nam1" value="">
	<input type="hidden" name="ngay2" value="">
	<input type="hidden" name="thang2" value="">
	<input type="hidden" name="nam2" value="">
	<input type="hidden" name="ViewType" value="">
	<input type="hidden" name="Username" value="">
</form>
<SCRIPT LANGUAGE=JavaScript>
<!--
 function view(ViewType,Username)
 {
 	document.fView.ngay1.value=document.fThongke.Ngay1.value;
 	document.fView.thang1.value=document.fThongke.Thang1.value;
 	document.fView.nam1.value=document.fThongke.Nam1.value;
 	document.fView.ngay2.value=document.fThongke.Ngay2.value;
 	document.fView.thang2.value=document.fThongke.Thang2.value;
 	document.fView.nam2.value=document.fThongke.Nam2.value;
 	document.fView.ViewType.value=ViewType;
 	document.fView.Username.value=Username;
 	document.fView.submit();
 }
// -->
</SCRIPT>

<table width="998"" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#000000">
  <tr align="center" valign="middle" bgcolor="#FFFFFF"> 
    <td width="5%" rowspan="2"><font size="2" face="Arial, Helvetica, sans-serif"><strong>TT</strong></font></td>
    <td width="15%" rowspan="2"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Biên tập viên</strong></font></td>
    <td width="40%" align="center" colspan="3"><font size="2" face="Arial, Helvetica, sans-serif"><b>Tin Text</b></font></td>
    <td width="20%" align="center" colspan="2"><font size="2" face="Arial, Helvetica, sans-serif"><b>Tin Audio Video</b></font></td>
    <td width="20%" align="center" colspan="2"><font size="2" face="Arial, Helvetica, sans-serif"><b>Tổng 2 loại</b></font></td>
  </tr>
  <tr align="center" valign="middle" bgcolor="#FFFFFF"> 
    <td width="20%" align="center"><font size="2" face="Arial, Helvetica, sans-serif">Số lượng<br>
      Đưa lên mạng</font></td>
    <td width="10%" align="center"><font size="2" face="Arial, Helvetica, sans-serif">Lượt xem</font></td>
    <td width="10%" align="center"><font size="2" face="Arial, Helvetica, sans-serif">Số bài<br>phản hồi</font></td>
    <td width="10%" align="center"><font size="2" face="Arial, Helvetica, sans-serif">Số lượng</font></td>
    <td width="10%" align="center"><font size="2" face="Arial, Helvetica, sans-serif">Lượt xem</font></td>
    <td width="10%" align="center"><font size="2" face="Arial, Helvetica, sans-serif">Số lượng</font></td>
    <td width="10%" align="center"><font size="2" face="Arial, Helvetica, sans-serif">Lượt xem</font></td>
  </tr>
  <%
  Dim rsNews
  Set rsNews=Server.CreateObject("ADODB.Recordset")
  Column_News=0
  Column_NewsCount=0
  Column_Comment=0
  Column_Av=0
  Column_AvCount=0
  Column_News_Av=0
  Column_News_Av_Count=0
  STT=0
  Do while not rsUser.eof 
  	sqlNews="SELECT	COUNT(NewsID) AS News_Total, Creator, SUM(NewsCount) AS NewsCount_Total " &_
		"FROM   V_News_Thongke " &_
		"WHERE (DATEDIFF(dd, CreationDate, '" & FromDate & "') <= 0) AND (DATEDIFF(dd, CreationDate, '" & ToDate & "') >= 0) " &_
		"GROUP BY Creator " &_
		"HAVING	(Creator = '" & rsUser("Username") & "')"
	rsNews.open sqlNews,con,3
		if IsNumeric(rsNews("News_Total")) then
			News_Total=Clng(rsNews("News_Total"))
		else
			News_Total=0
		end if
		if IsNumeric(rsNews("NewsCount_Total")) then
			NewsCount_Total=CLng(rsNews("NewsCount_Total"))
		else
			NewsCount_Total=0
		end if
	rsNews.close
	
	sqlComment="SELECT Count(CommentId) as Comment_Total " &_
		"FROM (	SELECT  nc.CommentID" &_
		"		FROM	V_News_Thongke v INNER JOIN " &_
        "				NewsComment nc ON v.NewsID = nc.NewsId " &_
		"		WHERE	(v.Creator = '" & rsUser("Username") & "') AND (DATEDIFF(dd, CreationDate, '" & FromDate & "') <= 0) AND (DATEDIFF(dd, CreationDate, '" & ToDate & "') >= 0) " &_
		"		GROUP BY nc.CommentID ) as View1"
	rsNews.open sqlComment,con,3
		if IsNumeric(rsNews("Comment_Total")) then
			Comment_Total=Clng(rsNews("Comment_Total"))
		else
			Comment_Total=0
		end if
	rsNews.close
	
	sqlAv="SELECT	Av_Creator, COUNT(Av_id) AS AV_Total, SUM(Av_Count) AS Av_Count_Total " &_
		"FROM	AudioVideo " &_
		"WHERE (DATEDIFF(dd, Av_CreationDate, '" & FromDate & "') <= 0) AND (DATEDIFF(dd, Av_CreationDate, '" & ToDate & "') >= 0) " &_
		"GROUP BY Av_Creator " &_
		"HAVING	(Av_Creator = '" & rsUser("Username") & "')"
	rsNews.open sqlAv,con,3
		if IsNumeric(rsNews("AV_Total")) then
			AV_Total=Clng(rsNews("AV_Total"))
		else
			AV_Total=0
		end if
		if IsNumeric(rsNews("Av_Count_Total")) then
			Av_Count_Total=Clng(rsNews("Av_Count_Total"))
		else
			Av_Count_Total=0
		end if
	rsNews.close
	News_Av=News_Total+AV_Total
	News_Av_Count=NewsCount_Total + Av_Count_Total
	if News_Av>0 then
		Column_News=Column_News + News_Total
		Column_NewsCount=Column_NewsCount + NewsCount_Total
		Column_Comment=Column_Comment + Comment_Total
		Column_Av=Column_Av + Av_Total
		Column_AvCount= Column_AvCount + Av_Count_Total
		Column_News_Av=Column_News_Av + News_Av
		Column_News_Av_Count=Column_News_Av_Count + News_Av_Count
		
		STT=STT+1
  		if bg_color="#E6E8E9" then
  			bg_color="#FFFFFF"
  		else
  			bg_color="#E6E8E9"
  		end if
  %>
  <tr bgcolor="<%=bg_color%>"> 
    <td align="right" valign="top"><font size="2" face="Arial, Helvetica, sans-serif"><%=STT%>.&nbsp;</font></td>
    <td><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;<%=rsUser("UserName")%></font></td>
    <td align="right"><font size="2" face="Arial, Helvetica, sans-serif"><a href="javascript: view('Online','<%=rsUser("UserName")%>');"><%=FormatNumber(News_Total,0)%></a>&nbsp;&nbsp;</font></td>
    <td align="right"><font size="2" face="Arial, Helvetica, sans-serif"><%=FormatNumber(NewsCount_Total,0)%>&nbsp;&nbsp;</font></td>
    <td align="right"><font size="2" face="Arial, Helvetica, sans-serif"><%=FormatNumber(Comment_Total,0)%>&nbsp;&nbsp;</font></td>
    <td align="right"><font size="2" face="Arial, Helvetica, sans-serif"><%=FormatNumber(AV_Total,0)%>&nbsp;&nbsp;</font></td>
    <td align="right"><font size="2" face="Arial, Helvetica, sans-serif"><%=FormatNumber(Av_Count_Total,0)%>&nbsp;&nbsp;</font></td>
    <td align="right"><font size="2" face="Arial, Helvetica, sans-serif"><%=FormatNumber(News_Av,0)%>&nbsp;&nbsp;</font></td>
    <td align="right"><font size="2" face="Arial, Helvetica, sans-serif"><%=FormatNumber(News_Av_Count,0)%>&nbsp;&nbsp;</font></td>
  </tr>
 <%	end if 'if News_Av>0 then
 	rsUser.movenext
 Loop
 %>
 <tr bgcolor="<%=bg_color%>">
    <td align="right" colspan="2"><font size="2" face="Arial, Helvetica, sans-serif"><b>TỔNG CỘNG &nbsp;</b></font></td>
    <td align="right"><font size="3" face="Arial, Helvetica, sans-serif"><b><%=FormatNumber(Column_News,0)%></b>&nbsp;&nbsp;</font></td>
    <td align="right"><font size="2" face="Arial, Helvetica, sans-serif"><b><%=FormatNumber(Column_NewsCount,0)%></b>&nbsp;&nbsp;</font></td>
    <td align="right"><font size="2" face="Arial, Helvetica, sans-serif"><b><%=FormatNumber(Column_Comment,0)%></b>&nbsp;&nbsp;</font></td>
    <td align="right"><font size="2" face="Arial, Helvetica, sans-serif"><b><%=FormatNumber(Column_Av,0)%></b>&nbsp;&nbsp;</font></td>
    <td align="right"><font size="2" face="Arial, Helvetica, sans-serif"><b><%=FormatNumber(Column_AvCount,0)%></b>&nbsp;&nbsp;</font></td>
    <td align="right"><font size="2" face="Arial, Helvetica, sans-serif"><b><%=FormatNumber(Column_News_Av,0)%></b>&nbsp;&nbsp;</font></td>
    <td align="right"><font size="2" face="Arial, Helvetica, sans-serif"><b><%=FormatNumber(Column_News_Av_Count,0)%></b>&nbsp;&nbsp;</font></td>
  </tr>
</table>
<%
  	Set rsNews=nothing
	rsUser.close
	set rsUser=nothing
'END IF 'IF Request.form("action")="Search" THEN
%>
</div>
</div>
    </div>
<%Call Footer()%>
</body>
</html>