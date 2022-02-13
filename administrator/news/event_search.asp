<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
if request.Form("action")="Search" then
	textkey=Trim(request.Form("textkey"))
	textkey=Replace(textkey,"'","''")

	if textkey="" then
		response.redirect ("/administrator/")
		response.End()
	end if
	
	
	if textkey="$Tất cả$" then
		sql1=GetSQL_For_Search(session("LstCat"),session("LstRole"),session("user"),"NONE")
		sql="SELECT *"
		sql=sql & " FROM event e, NewsDistribution d, NewsCategory c"
		sql=sql & " WHERE n.NewsId=d.NewsId and d.CategoryId=c.CategoryId and " & sql1
		sql=sql & " ORDER BY n.NewsId desc"
	else
		sql1=GetSQL_For_Search(session("LstCat"),session("LstRole"),session("user"),Cat)
		sql="SELECT n.NewsId, n.Title,n.Creator,n.CreationDate,n.LastEditor,n.LastEditedDate,"
		sql=sql & "n.StatusId,d.CategoryId,c.CategoryName"
		sql=sql & " FROM News n, NewsDistribution d, NewsCategory c"
		sql=sql & " WHERE n.NewsId=d.NewsId and d.CategoryId=c.CategoryId and "
		if Cat="Marked" or Cat="Edit" or Cat="Waiting" then
			sql=sql & sql1
		elseif Cat<>"PublicationNo" and Cat<>"NewsID" then
			sql=sql & "n." & Cat & " like N'%" & textkey  & "%' and " & sql1
		else
			sql=sql & "n." & Cat & "=" & textkey  & " and " & sql1
		end if
		sql=sql & " ORDER BY n.NewsId desc"
	end if
end if
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
	Title_This_Page="Quản lý -> Tìm kiếm tin &SP"
	Call header()
	
	
	
%>
<FORM action="<%=Request.ServerVariables("script_name")%>" method="post" name="fSearch">
<table align="center" cellpadding="0" cellspacing="0" width="770">
	<tr> 
		<td align="right" valign="middle" width="100%">
			<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Tìm kiếm:&nbsp;</strong></font> 
		</td>
		<td align="left" valign="middle" width="1">
			<input name="textkey" type="text" id="textkey" value="<%if textkey<>"" then response.Write(textkey) else response.Write("Unicode font")%>" size="25" maxlength="25" onFocus="javascript: this.value='';">
		</td>
		<td align="left" valign="middle" width="1">
      		<a href="#" onClick="javascript: checkme(fSearch.Cat.value,fSearch.textkey.value);"><img name="ButtonSearch" id="ButtonSearch" align="absmiddle" src="/administrator/images/search_bt.gif" width="23" height="23" border="0"></a>
			<input type="hidden" name="action" value="Search">
    	</td>
	</tr>
</table>
</form>
<SCRIPT LANGUAGE=JavaScript>
<!--
 function checkme(thisCatvalue,thisTextvalue)
 {
 	if (thisCatvalue==0)
	{
		alert("Bạn chưa chọn phạm vi tìm kiếm");
		fSearch.Cat.focus();
		return false;
	}
	if ((thisTextvalue.length==0) && (thisCatvalue!="Marked") && (thisCatvalue!="Edit") && (thisCatvalue!="Waiting"))
	{
		alert("Bạn chưa nhập khóa tìm kiếm");
		fSearch.textkey.focus();
		return false;
	}
	if ( (!IsNumeric(thisTextvalue)) && ((thisCatvalue=="NewsID") || (thisCatvalue=="PublicationNo")))
	{
		alert("Bạn chưa nhập sai định dạng tìm kiếm");
		fSearch.textkey.value="";
		fSearch.textkey.focus();
		return false;
	}
	document.fSearch.submit();
 }
// -->
</SCRIPT>
<%if request.Form("action")="Search" then
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	NEWS_PER_PAGE=20
	PAGES_PER_BOOK=7
	rs.PageSize = NEWS_PER_PAGE
	'response.Write(sql)
	'response.end
	rs.open sql,con,1
	
	if not rs.eof then
		if request.Form("page")<>"" and isnumeric(Request.Form("page")) then
			page=Cint(request.Form("page"))
		else
			page=1
		end if
		rs.AbsolutePage = CLng(page)
		stt=(page-1)* rs.pageSize + 1
		i=0
%>
<table width="770" align="center" cellpadding="0" cellspacing="1" bordercolor="#000000" bgcolor="#000000">
  <tr bgcolor="#FFFFFF"> 
    <td> <div align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">TT</font></strong></div></td>
    <td> <div align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Tiêu 
        đề tin</font></strong></div></td>
    <td> <div align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Chuyên 
        mục</font></strong></div></td>
    <td> <div align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Trạng 
        thái</font></strong></div></td>
    <td> <div align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Xử 
        lý</font></strong></div></td>
  </tr>
<%Do while not rs.eof and i<rs.pagesize
	i=i+1
%>
  <tr bgcolor="#FFFFFF"> 
    <td align="right" valign="top"><font size="2" face="Arial, Helvetica, sans-serif"><%=stt%>.</font></td>
    <td><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;<a href="javascript: winpopup('/administrator/news/news_view.asp','<%=rs("NewsId")%>&CatId=<%=rs("CategoryId")%>',600,400);" style="text-decoration: none"><%=rs("Title")%></a><br>
      <font size="1">&nbsp;(Tạo bởi: <%=rs("Creator")%>-<%=Hour(ConvertTime(rs("CreationDate")))%>h<%=Minute(ConvertTime(rs("CreationDate")))%>&quot;&nbsp;<%=Day(ConvertTime(rs("CreationDate")))%>/<%=Month(ConvertTime(rs("CreationDate")))%>/<%=Year(ConvertTime(rs("CreationDate")))%>, 
	  Sửa lần cuối: <%=rs("LastEditor")%>-<%=Hour(ConvertTime(rs("LastEditedDate")))%>h<%=Minute(ConvertTime(rs("LastEditedDate")))%>&quot;&nbsp;<%=Day(ConvertTime(rs("LastEditedDate")))%>/<%=Month(ConvertTime(rs("LastEditedDate")))%>/<%=Year(ConvertTime(rs("LastEditedDate")))%>)</font>
      </font></td>
    <td align="center" valign="middle"><font size="2" face="Arial, Helvetica, sans-serif"><%=GetListParentCatNameOfCatId(rs("CategoryId"))%></font></td>
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
    <td align="center" valign="middle">
		<a href="javascript: winpopup('/administrator/news/news_viewnote.asp','<%=rs("NewsId")%>&CatId=<%=rs("CategoryId")%>',300,300);"><img src="../images/help-book.gif" width="15" height="15" border="0" align="absmiddle" title="Xem các ghi chú"></a>
    	<a href="news_edit.asp?newsid=<%=rs("NewsId")%>&catid=<%=rs("Categoryid")%>"><img src="../images/icon_edit_topic.gif" width="15" height="15" border="0" align="absmiddle" title="Sửa"></a> 
      	<a href="javascript: winpopup('/administrator/news/news_delete.asp','<%=rs("NewsId")%>&CatId=<%=rs("CategoryId")%>',400,150);"><img src="../images/icon_closed_topic.gif" width="15" height="15" border="0" align="absmiddle" title="Xóa"></a> 
    </td>
  </tr>
<%stt=stt+1
rs.movenext
Loop%>
	<form action="<%=Request.ServerVariables("Script_name")%>" method="post" name="fSearch2">
		<input type="hidden" name="textkey" value="<%=textkey%>">
		<input type="hidden" name="Cat" value="<%=Cat%>">
		<input type="hidden" name="page">
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
									if page>Cint(PAGES_PER_BOOK/2) then
										Minpage=page-Cint(PAGES_PER_BOOK/2)+1
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
<%	end if 'if not rs.eof then
	rs.close
	set rs=nothing
else
	response.Write("<p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p>")
end if%>
<%Call Footer()%>
</body>
</html>