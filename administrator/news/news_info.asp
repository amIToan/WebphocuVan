<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<HTML>
	<HEAD>
		<TITLE><%=PAGE_TITLE%></TITLE>
		<META http-equiv=Content-Type content="text/html; charset=utf-8">
		<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
	</HEAD>
<BODY leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
	Call header()
	
	Call TitlePage("Th&#244;ng b&#225;o!")
%>
<table width="100%" border="0" cellspacing="2" cellpadding="2" align="center">
	<tr> 
		<td height="40">&nbsp;</td>
	</tr>
	<tr align="center" valign="middle"> 
		<td align="center" valign="middle">
			<br><font size="4" face="Verdana, Arial, Helvetica, sans-serif">
				Bạn không được sửa tin.
			</font><br><br>
			<%
			NewsId=Request.QueryString("NewsId")
			CatId=Request.QueryString("CatId")
			if not IsNumeric(NewsId) or not IsNumeric(CatId) then
				response.Redirect("/administrator/")
				response.End()
			else
				NewsId=CLng(NewsId)
				CatId=CLng(CatId)
			end if
			Dim rs
			set rs=server.CreateObject("ADODB.Recordset")
			sql="SELECT top 1 d.CategoryId, d.NewsId, u.Username, u.UserEmail"
			sql=sql & " FROM NewsDistribution d,[User] u"
			sql=sql & " WHERE d.NewsId=" & NewsId & " and (d.StatusId='apap' or d.StatusId='adad') and (d.Approver=u.Username or d.Administrator=u.Username) and d.CategoryId<>" & CatId
			'response.write sql
			rs.open sql,con,1
			if rs.eof then
				rs.close
				set rs=nothing
				response.end()
			end if
			
			%>
			<font size="2" face="Verdana, Arial, Helvetica, sans-serif">
				Vì tin đã được đưa lên mạng ở một chuyên mục khác (<a href="javascript: winpopup('/administrator/news/news_view.asp','<%=rs("NewsId")%>&CatId=<%=rs("CategoryId")%>',600,400);">xem</a>).<br>
				Nếu như bạn chỉnh sửa tin -> thay đổi thông tin đã đưa lên mạng.<br>
				Bạn có thể gửi yêu cầu đến <%=rs("Username")%> (Email: <a href="mailto: <%=rs("UserEmail")%>"><%=rs("UserEmail")%></a>) để sửa tin.<br>
			</font>
			<%rs.close
			set rs=nothing%>
		</td>
	</tr>
	<tr> 
		<td height="80">&nbsp;</td>
	</tr>
</table>
<%Call Footer()%>
</BODY>
</HTML>