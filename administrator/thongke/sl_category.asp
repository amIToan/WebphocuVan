<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
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
   ' Response.Write FromDate&" - "&ToDate
%>
<html>
<head>
<title><%=PAGE_TITLE%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
    <LINK href="/administrator/inc/admstyle.css" type=text/css rel=stylesheet>
    <link href="../../css/styles.css" rel="stylesheet" type="text/css">
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
	Title_This_Page="Thống kê -> SL tin theo Chuyên mục"
	Call header()	
%>
<div class="container-fluid">
    <div class="col-md-2" style="background:#001e33;"><%call MenuVertical(0) %></div>
    <div class="col-md-10  style="background:#001e33">
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
    function order(OrderType) {
        if (!checkme())
            return;
        document.fThongke.OrderType.value = OrderType;
        document.fThongke.submit();
    }
    function checkme() {
        if (document.fThongke.Ngay1.value == 0) {
            alert("Bạn chưa chọn ngày!");
            document.fThongke.Ngay1.focus();
            return false;
        }
        if (document.fThongke.Thang1.value == 0) {
            alert("Bạn chưa chọn tháng!");
            document.fThongke.Thang1.focus();
            return false;
        }
        if (document.fThongke.Nam1.value == 0) {
            alert("Bạn chưa chọn năm!");
            document.fThongke.Nam1.focus();
            return false;
        }
        if (document.fThongke.Ngay2.value == 0) {
            alert("Bạn chưa chọn ngày!");
            document.fThongke.Ngay2.focus();
            return false;
        }
        if (document.fThongke.Thang2.value == 0) {
            alert("Bạn chưa chọn tháng!");
            document.fThongke.Thang2.focus();
            return false;
        }
        if (document.fThongke.Nam2.value == 0) {
            alert("Bạn chưa chọn năm!");
            document.fThongke.Nam2.focus();
            return false;
        }
        return true;
    }
// -->
</SCRIPT>
<%
'IF Request.form("action")="Search" And IsDate(ToDate) AND IsDate(FromDate) THEN
	'FromDate=FormatDateTime(FromDate)
	'ToDate=FormatDateTime(ToDate)
  
	Dim rsCat
	Set rsCat=Server.CreateObject("ADODB.Recordset")
	
	'Lấy danh sách Chuyên mục được quyền hiển thị
	if Trim(session("LstCat"))="0" then
		sqlCat="SELECT	CategoryID, CategoryName, CategoryLevel, YoungestChildren, CategoryLoai " &_
			"FROM	NewsCategory " &_
			"ORDER BY LanguageId DESC, CategoryOrder"
	else
		sqlCat="SELECT	CategoryID, CategoryName, CategoryLevel, YoungestChildren " &_
			"FROM	NewsCategory "
		strCat=GetListChildrenOfListCat(session("LstCat")) & " " & GetListParentOfListCat(session("LstCat"))
		if strCat<>"" then
			ArrCat=Split(" " & strCat & " ")
			j=0
			for i=1 to UBound(ArrCat)
				if IsNumeric(ArrCat(i)) then
					j=j+1
					if j=1 then
						sqlCat=sqlCat & "Where CategoryId=" & ArrCat(i)
					else
						sqlCat=sqlCat & " or CategoryId=" & ArrCat(i)
					end if
				end if
			next
		end if 'if strCat<>"" then
		sqlCat=sqlCat & " ORDER BY LanguageId DESC, CategoryOrder"
	end if 'if Trim(session("LstCat"))="0" then
	
	'response.write sqlCat
	rsCat.open sqlCat,con,3

%>

<table width="998" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#000000" class="CTxtContent w3-table w3-table-all w3-round w3-margin">
  <tr align="center" valign="middle" bgcolor="#FFFFFF"> 
    <td width="4%"><font size="2" face="Arial, Helvetica, sans-serif"><strong>TT</strong></font></td>
    <td width="40%"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Tên chuyên mục</strong></font></td>
    <td width="19%" align="center"><font size="2" face="Arial, Helvetica, sans-serif">Số 
      lượng tin đang<br>
      Đưa lên mạng</font></td>
    <td width="13%" align="center"><font size="2" face="Arial, Helvetica, sans-serif">Lượt 
      xem<br>
      Tổng cộng</font></td>
    <td width="17%" align="center"><font size="2" face="Arial, Helvetica, sans-serif">Số 
      bài<br>
      độc giả phản hồi</font></td>
  </tr>
  <%
  'Số tin Online
  NewsOnline_Total=0 'Tổng số
  'Lượt xem tin
  NewsCount_Total=0 'Tổng số
  'Phản hồi của độc giả
  Reply_Total=0 'Tổng số

  STT=0
  Dim rsCount
  Set rsCount=Server.CreateObject("ADODB.Recordset")
  HTML=""
  sHTML=""
  Do while not rsCat.eof
  	STT=STT+1
  	if bg_color="#E6E8E9" then
  		bg_color="#FFFFFF"
  	else
  		bg_color="#E6E8E9"
  	end if
  	
  	if rsCat("CategoryId")=61 then
  	'Đếm thống kê cho Audio_Video
  		sqlCount_AV="SELECT	 COUNT(Av_id) AS Av_Total, SUM(Av_Count) AS Av_Count_Total " &_
			"FROM	AudioVideo " &_
			"WHERE (DATEDIFF(dd, Av_CreationDate, '" & FromDate & "') <= 0) AND (DATEDIFF(dd, Av_CreationDate, '" & ToDate & "') >= 0)"
		
		rsCount.open sqlCount_AV,con,3
    		Av_Total=Clng(rsCount("Av_Total"))
    		if IsNumeric(rsCount("Av_Count_Total")) then
  				Av_Count_Total=Clng(rsCount("Av_Count_Total"))
  			else
  				Av_Count_Total=0
  			end if
  		rsCount.close
		Reply_Total=0
		
  	elseif Clng(rsCat("CategoryLevel"))=1 then
  	'Count News
  	sqlCount_News="SELECT COUNT(NewsId) as NewsOnline_Total, SUM(NewsCount) as NewsCount_Total " &_
		"FROM (	SELECT	NewsID, COUNT(NewsID) AS Num_News, AVG(NewsCount) AS NewsCount " &_
		"		FROM         V_News_Thongke " &_
		"		WHERE (CategoryId=" & rsCat("CategoryId") & " Or ParentCategoryID=" & rsCat("CategoryId") & ") AND (DATEDIFF(dd, CreationDate, '" & FromDate & "') <= 0) AND (DATEDIFF(dd, CreationDate, '" & ToDate & "') >= 0) " &_
		"		GROUP BY NewsID " &_
		") as View1"
	'Count Reply Comment
	sqlCount_Reply="Select Count(CommentId) as Reply_Total " &_
		"FROM ( SELECT	nc.CommentID, COUNT(nc.CommentID) AS Reply_Total " &_
		"		FROM	NewsComment nc INNER JOIN " &_
        "				V_News_Thongke v ON nc.NewsId = v.NewsID " &_
		"		WHERE	(nc.SubjectId = 0) AND (v.CategoryId=" & rsCat("CategoryId") & " Or v.ParentCategoryID=" & rsCat("CategoryId") & ") AND (DATEDIFF(dd, v.CreationDate, '" & FromDate & "') <= 0) AND (DATEDIFF(dd, v.CreationDate, '" & ToDate & "') >= 0) " &_
		"		GROUP BY nc.CommentID " &_
		"	   ) As View1"
      'Response.Write sqlCount_News
		rsCount.open sqlCount_News,con,3
    		NewsOnline_Total=Clng(rsCount("NewsOnline_Total"))
    		if IsNumeric(rsCount("NewsCount_Total")) then
  				NewsCount_Total=Clng(rsCount("NewsCount_Total"))
  			else
  				NewsCount_Total=0
  			end if
  		rsCount.close
		rsCount.open sqlCount_Reply,con,3
			Reply_Total=Clng(rsCount("Reply_Total"))
		rsCount.close
	else
  	'Count News
  	sqlCount_News="SELECT COUNT(NewsId) as NewsOnline_Total, SUM(NewsCount) as NewsCount_Total " &_
		"FROM (	SELECT	NewsID, COUNT(NewsID) AS Num_News, AVG(NewsCount) AS NewsCount " &_
		"		FROM         V_News_Thongke " &_
		"		WHERE (CategoryId=" & rsCat("CategoryId") & ") AND (DATEDIFF(dd, CreationDate, '" & FromDate & "') <= 0) AND (DATEDIFF(dd, CreationDate, '" & ToDate & "') >= 0) " &_
		"		GROUP BY NewsID " &_
		") as View1"
	
	'Count Reply Comment
	sqlCount_Reply="Select Count(CommentId) as Reply_Total " &_
		"FROM ( SELECT	nc.CommentID, COUNT(nc.CommentID) AS Reply_Total " &_
		"		FROM	NewsComment nc INNER JOIN " &_
        "				V_News_Thongke v ON nc.NewsId = v.NewsID " &_
		"		WHERE	(nc.SubjectId = 0) AND (v.CategoryId=" & rsCat("CategoryId") & ") AND (DATEDIFF(dd, v.CreationDate, '" & FromDate & "') <= 0) AND (DATEDIFF(dd, v.CreationDate, '" & ToDate & "') >= 0) " &_
		"		GROUP BY nc.CommentID " &_
		"	   ) As View1"
		rsCount.open sqlCount_News,con,3
    		NewsOnline_Total=Clng(rsCount("NewsOnline_Total"))
    		if IsNumeric(rsCount("NewsCount_Total")) then
  				NewsCount_Total=Clng(rsCount("NewsCount_Total"))
  			else
  				NewsCount_Total=0
  			end if
  		rsCount.close
		rsCount.open sqlCount_Reply,con,3
			Reply_Total=Clng(rsCount("Reply_Total"))
		rsCount.close
	end if 'if rsCat("CategoryId")=61 then
	
	'Hiển thị HTML
	if Clng(rsCat("Categoryid"))=61 then
%>		
<tr bgcolor="<%=bg_color%>">
	<td><%=STT%>.</td>
	<td>
	<b>&nbsp;&nbsp;&nbsp;-&nbsp;<a href="sl_prints.asp?CatID=<%=rsCat("CategoryId")%>" target="_blank"><%=rsCat("CategoryName")%></a></b>
	<%if rsCat("CategoryLoai")= 3 then%>
		<div align="right" class="CSubTitle"><a href="sl_printsEdit.asp?CatID=<%=rsCat("CategoryId")%>" target="_blank">Sửa chiết khấu</a></div>
	<%end if%>
	</td>
	<td><%=CStr(FormatNumber(NewsOnline_Total,0))%></td>
	<td><%=CStr(FormatNumber(NewsCount_Total,0))%></td>
	<td><%=CStr(FormatNumber(Reply_Total,0))%></td>
</tr>
<%	
	elseif Clng(rsCat("CategoryLevel"))=1 then
%>		
<tr bgcolor="<%=bg_color%>">
	<td><%=STT%>.</td>
	<td><b>&nbsp;&#8226;&nbsp;<a href="sl_prints.asp?CatID=<%=rsCat("CategoryId")%>" target="_blank"><%=rsCat("CategoryName")%></a></b>
	<%if rsCat("CategoryLoai")= 3 then%>
		<div align="right"><a href="sl_printsEdit.asp?CatID=<%=rsCat("CategoryId")%>" target="_blank">Sửa chiết khấu</a></div>
	<%end if%>	
	</td>
	<td><%=CStr(FormatNumber(NewsOnline_Total,0))%></td>
	<td><%=CStr(FormatNumber(NewsCount_Total,0))%></td>
	<td><%=CStr(FormatNumber(Reply_Total,0))%></td>
</tr>
<%
  	else

 %>		
<tr bgcolor="<%=bg_color%>">
	<td><%=STT%>.</td>
	<td>
	<b>&nbsp;&nbsp;&nbsp;-&nbsp;<a href="sl_prints.asp?CatID=<%=rsCat("CategoryId")%>" target="_blank"><%=rsCat("CategoryName")%></a></b>
	<%if rsCat("CategoryLoai")= 3 or rsCat("CategoryLoai")= 7 or rsCat("CategoryLoai")= 10 then%>
		<div align="right"><a href="sl_printsEdit.asp?CatID=<%=rsCat("CategoryId")%>" target="_blank">Sửa chiết khấu</a></div>
	<%end if%>
	</td>
	<td><%=CStr(FormatNumber(NewsOnline_Total,0))%></td>
	<td><%=CStr(FormatNumber(NewsCount_Total,0))%></td>
	<td><%=CStr(FormatNumber(Reply_Total,0))%></td>
</tr>
<%
 	end if
 rsCat.movenext
 Loop

'Đếm tổng số
  	'Count News
  	sqlCount_News="SELECT COUNT(NewsId) as NewsOnline_Total, SUM(NewsCount) as NewsCount_Total " &_
		"FROM (	SELECT	NewsID, COUNT(NewsID) AS Num_News, AVG(NewsCount) AS NewsCount " &_
		"		FROM         V_News_Thongke " &_
		"		WHERE (DATEDIFF(dd, CreationDate, '" & FromDate & "') <= 0) AND (DATEDIFF(dd, CreationDate, '" & ToDate & "') >= 0) " &_
		"		GROUP BY NewsID " &_
		") as View1"
	
	'Count Reply Comment
	sqlCount_Reply="Select Count(CommentId) as Reply_Total " &_
		"FROM ( SELECT	nc.CommentID, COUNT(nc.CommentID) AS Reply_Total " &_
		"		FROM	NewsComment nc INNER JOIN " &_
        "				V_News_Thongke v ON nc.NewsId = v.NewsID " &_
		"		WHERE	(nc.SubjectId = 0) AND (DATEDIFF(dd, v.CreationDate, '" & FromDate & "') <= 0) AND (DATEDIFF(dd, v.CreationDate, '" & ToDate & "') >= 0) " &_
		"		GROUP BY nc.CommentID " &_
		"	   ) As View1"

    rsCount.open sqlCount_News,con,3
    	NewsOnline_Total=Clng(rsCount("NewsOnline_Total"))
    	if IsNumeric(rsCount("NewsCount_Total")) then
  			NewsCount_Total=Clng(rsCount("NewsCount_Total"))
  		else
  			NewsCount_Total=0
  		end if
  	rsCount.close
  	
	rsCount.open sqlCount_Reply,con,3
		Reply_Total=Clng(rsCount("Reply_Total"))
 	rsCount.close
 	
 Set rsCount=nothing
 rsCat.close
 set rsCat=nothing
 %>
 <tr bgcolor="<%=bg_color%>">
    <td align="right" colspan="2"><font size="2" face="Arial, Helvetica, sans-serif"><b>TỔNG CỘNG &nbsp;</b></font></td>
    <td align="right"><font size="3" face="Arial, Helvetica, sans-serif"><b><%=FormatNumber(NewsOnline_Total + Av_Total,0)%>&nbsp;&nbsp;</b></font></td>
    <td align="right"><font size="2" face="Arial, Helvetica, sans-serif"><b><%=FormatNumber(NewsCount_Total + Av_Count_Total ,0)%>&nbsp;&nbsp;</b></font></td>
    <td align="right"><font size="2" face="Arial, Helvetica, sans-serif"><b><%=FormatNumber(Reply_Total,0)%>&nbsp;&nbsp;</b></font></td>
  </tr>
</table>

<%
'END IF 'IF Request.form("action")="Search" THEN
%>
</div>
    </div>
        </div>
<%Call Footer()%>
</body>
</html>