<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
f_permission = administrator(false,session("user"),"m_report")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
	Username=ReplaceHTMLToText(Request.form("Username"))
	if session("LstRole")="0ad" then
		if session("User")<> "admin" and CheckUserExist(Username)=0 then
		'Không tồn tại User tìm kiếm
			response.end
		end if
	else
		if UCase(UserName)<>UCase(Session("User")) then
		'Nếu không có quyền admin, chỉ được phép xem thông tin của chính mình
			response.end
		end if
	end if
	Ngay1=GetNumeric(Request.form("Ngay1"),0)
	Thang1=GetNumeric(Request.form("Thang1"),0)
	Nam1=GetNumeric(Request.form("Nam1"),0)
	Ngay2=GetNumeric(Request.form("Ngay2"),0)
	Thang2=GetNumeric(Request.form("Thang2"),0)
	Nam2=GetNumeric(Request.form("Nam2"),0)
	ViewType=ReplaceHTMLToText(Request.form("ViewType"))
%>
<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
<LINK href="/administrator/inc/admstyle.css" type=text/css rel=stylesheet>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%	Title_This_Page="Th&#7889;ng k&#234; -> SL tin theo Biên tập viên"
	Call header()
	Call Menu()

	
%>
<table width="770" align="center" cellpadding="0" cellspacing="1" border="0">
  <tr> 
    <td align="center"><br><font size="2" face="Arial, Helvetica, sans-serif">Danh sách tin đang
    <%Select case ViewType
    	case "Online"
    		Response.write "<b>đưa lên mạng</b>"
    End Select%>
   	của <b><%=Username%></b> từ ngày <b><%=Ngay1%>/<%=Thang1%>/<%=Nam1%></b> đến ngày <b><%=Ngay2%>/<%=Thang2%>/<%=Nam2%></b></font><br><br></td>
  </tr>
  <%
  	FromDate=Thang1 & "/" & Ngay1 & "/" & Nam1
  	ToDate=Thang2 & "/" & Ngay2 & "/" & Nam2
  	
  	FromDate=FormatDatetime(FromDate)
	ToDate=FormatDatetime(ToDate)
	
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	
	sql="SELECT Sum(n.NewsCount) as NewsCount_Total"
		sql=sql & " FROM News n"
		sql=sql & " WHERE  n.Creator='" & Username & "'"
		sql=sql & " and (DATEDIFF(dd, n.CreationDate, '" & FromDate & "') <= 0) AND (DATEDIFF(dd, n.CreationDate, '" & ToDate & "') >= 0) "
		if ViewType="Online" then
			sql=sql & " and (n.StatusId='apap' or n.StatusId='adad')"
		end if
	
	rs.open sql,con,3
	NewsCount_Total=Clng(rs("NewsCount_Total"))
	rs.close
  %>
  <tr> 
    <td align="left"><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;&nbsp;Tổng cộng có <b><%=FormatNumber(NewsCount_Total,0)%></b> lượt xem</font></td>
  </tr>
</table>
<%
	sql="SELECT n.NewsId, n.Title,n.Creator,n.CreationDate,n.LastEditor,n.LastEditedDate,"
		sql=sql & "n.StatusId, n.NewsCount"
		sql=sql & " FROM News n"
		sql=sql & " WHERE  n.Creator='" & Username & "'"
		sql=sql & " and (DATEDIFF(dd, n.CreationDate, '" & FromDate & "') <= 0) AND (DATEDIFF(dd, n.CreationDate, '" & ToDate & "') >= 0) "
		if ViewType="Online" then
			sql=sql & " and (n.StatusId='apap' or n.StatusId='adad')"
		end if
		sql=sql & " ORDER BY n.NewsId desc"

	
	NEWS_PER_PAGE=50
	PAGES_PER_BOOK=15
	rs.PageSize = NEWS_PER_PAGE

	rs.open sql,con,1
	
	if not rs.eof then
		Dim rsReply
		Set rsReply=Server.CreateObject("ADODB.Recordset")
		
		if request.Form("page")<>"" and isnumeric(Request.Form("page")) then
			page=Clng(request.Form("page"))
		else
			page=1
		end if
		rs.AbsolutePage = CLng(page)
		stt=(page-1)* rs.pageSize + 1
		iSTT=0
%>
<table width="770" align="center" cellpadding="0" cellspacing="1" bordercolor="#000000" bgcolor="#000000">
  <tr bgcolor="#FFFFFF"> 
    <td><div align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">TT</font></strong></div></td>
    <td><div align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Tiêu 
        đề tin</font></strong></div></td>
    <td><div align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Trạng<br>
        thái</font></strong></div></td>
	<td><div align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Lượt<br>xem</font></strong></div></td>
	<td><div align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Phản<br>hồi</font></strong></div></td>
  </tr>
<%
  Do while not rs.eof and iSTT<rs.pagesize
	iSTT=iSTT+1
%>
  <tr bgcolor="#FFFFFF"> 
    <td align="right" valign="top"><font size="2" face="Arial, Helvetica, sans-serif"><%=stt%>.</font></td>
    <td><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;<a href="javascript: winpopup('/administrator/news/news_view.asp','<%=rs("NewsId")%>&CatId=1',600,400);" style="text-decoration: none"><%=rs("Title")%></a><br>
      <font size="1">&nbsp;(Tạo bởi: <%=rs("Creator")%>-<%=Hour(ConvertTime(rs("CreationDate")))%>h<%=Minute(ConvertTime(rs("CreationDate")))%>&quot;&nbsp;<%=Day(ConvertTime(rs("CreationDate")))%>/<%=Month(ConvertTime(rs("CreationDate")))%>/<%=Year(ConvertTime(rs("CreationDate")))%>, 
	  Sửa lần cuối: <%=rs("LastEditor")%>-<%=Hour(ConvertTime(rs("LastEditedDate")))%>h<%=Minute(ConvertTime(rs("LastEditedDate")))%>&quot;&nbsp;<%=Day(ConvertTime(rs("LastEditedDate")))%>/<%=Month(ConvertTime(rs("LastEditedDate")))%>/<%=Year(ConvertTime(rs("LastEditedDate")))%>)</font>
      </font></td>
    <td align="center" valign="middle"><font size="2" face="Arial, Helvetica, sans-serif">
		<%statusid=rs("statusid")
		if right(StatusId,2)="ma" then
			response.Write("Đánh dấu")
		elseif CompareRole(left(StatusId,2), right(StatusId,2))>0 then
			response.Write("Yêu cầu sửa")
		elseif CompareRole(left(StatusId,2), right(StatusId,2))<0 then
			response.Write("Chờ duyệt")
		elseif StatusId="apap" or StatusId="adad" then
			response.Write("Lên mạng")
		end if%>
	</font></td>
	<td align="center" valign="middle"><font size="2" face="Arial, Helvetica, sans-serif"><%=rs("NewsCount")%></font></td>
	<%
		sqlReply="SELECT	COUNT(ID) AS Dem " &_
				 "FROM	Y_KIEN " &_
        		 "WHERE NewsId = " & rs("NewsId")
		rsReply.open sqlReply,con,3
	%>
	<td align="center" valign="middle"><font size="2" face="Arial, Helvetica, sans-serif"><%=rsReply("Dem")%></font></td>
  </tr>
<%		rsReply.close
stt=stt+1
rs.movenext
Loop%>
  
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
		<input type="hidden" name="action" value="Search">
	</form>
	<script language="JavaScript">
		function fSearch2(page)
		{
			document.fSearch2.page.value=page;
			document.fSearch2.submit();
		}
	</script>
                            <tr align="center" bgcolor="#FFFFFF"> 
                              <td colspan="5">
							  <font size="1" face="Verdana, Arial, Helvetica, sans-serif">
							  	Trang: 
									<%
									if page>Clng(PAGES_PER_BOOK/2) then
										Minpage=page-Clng(PAGES_PER_BOOK/2)+1
									else
										MinPage=1
									end if
									
									Maxpage=Minpage+PAGES_PER_BOOK-1
									if Maxpage >rs.pagecount then
										Maxpage=rs.pagecount
									end if
									
									for i=Minpage to Maxpage 
										if i<>page then
										response.Write "<a href=""javascript: fSearch2(" & i & ");"">" & i & "</a>&nbsp;|&nbsp;"
									  else
									  	response.Write "<font color=""red""><b>" & i & "</b></font>&nbsp;|&nbsp;"
									  end if
									Next%>
							   </font>
							  </td>
                            </tr>
</table>
<%		set rsReply=nothing
	end if 'if not rs.eof then
	rs.close
	set rs=nothing
%>
<%Call Footer()%>
</body>
</html>