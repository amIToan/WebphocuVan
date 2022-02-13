<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%	
Call Authenticate("None")
%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
<LINK href="/administrator/inc/admstyle.css" type=text/css rel=stylesheet>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
	NewsId=Request.QueryString("param")
	CatId=Request.QueryString("CatId")
	
	if not IsNumeric(NewsId) or not IsNumeric(CatId) then
		response.Redirect("/administrator/")
		response.End()
	else
		NewsId=CLng(NewsId)
		CatId=CLng(CatId)
	end if
	
	'Call AuthenticateWithRole(CatId,Session("LstRole"),"ed")
	'Phải có quyền xem (Editor) trở lên mới có thể xem tin này
	'Ngăn không cho các User khác, không có quyền vào chuyên mục xem tin trong chuyên mục này
	
	sql="select n.SubTitle, n.Title,n.CreationDate, n.PictureId, n.PictureAlign, n.Description, n.body, n.Author, n.Source"
	sql=sql & " FROM News n,NewsDistribution d, NewsCategory c"
	sql=sql & " WHERE n.NewsId=d.NewsId and d.CategoryId=c.CategoryId and n.NewsId=" & NewsId '& " and c.CategoryId=" & CatId
	Dim rs

    Response.Write(sql)

	set rs=server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
	if rs.eof then
		rs.close
		set rs=nothing
		response.End
	end if
%>
<TABLE>
	<TBODY>
		<tr> 
          <td height="30" valign="top">
		  	<%if Trim(rs("SubTitle"))<>"" then%>
			   <font class="main_subtitle"><%=rs("SubTitle")%></font><br>
			<%End if%>
			<font size="2" face="arial"><strong><font color="856C34"><%=rs("Title")%></font></strong></font>
			<br><font class="News_Date"><%=Hour(ConvertTime(rs("CreationDate")))%>h<%=Minute(ConvertTime(rs("CreationDate")))%>'&nbsp;<%=Day(ConvertTime(rs("CreationDate")))%>/<%=Month(ConvertTime(rs("CreationDate")))%>/<%=Year(ConvertTime(rs("CreationDate")))%></font>
		  </td>
        </tr>
        <tr> 
          <td>
		  	<%if rs("Pictureid")<>0 then 
		  		Call ShowPicture(rs("Pictureid"),1,rs("PictureAlign"))
		  	End if%>
			<p align="justify"><font size="2" face="arial"><strong><em>
				<%=rs("Description")%>
			</em></strong></font> </p>
            <p align="justify"><font size="2" face="arial">
				<%


                    Response.Write("aaaaaaaaaaaa"&rs("body")& "ZZZZZZZZZZZZZ")

					content=Replace(rs("body"),"&shy;","")
					content=Replace(content,"&#39;","'")
					Response.Write content
				%>
			</font></p>
            <p align="right"><strong><font size="2" face="Arial, Helvetica, sans-serif">

			</font></strong></p>
            </td>
        </tr>
		<tr><td align="center"><font size="2" face="Arial, Helvetica, sans-serif"><a href="javascript: window.close();">Đóng cửa sổ</a></font></td></tr>
	</TBODY>
</TABLE>
<%rs.close
set rs=nothing%>
</body>
</html>