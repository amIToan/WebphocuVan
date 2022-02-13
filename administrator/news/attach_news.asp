<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/include/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<html >
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=PAGE_TITLE%></title>
<link href="../../css/styles.css" rel="stylesheet" type="text/css">
    <script src="../../Scripts/news.js"></script>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
	img	="../../images/icons/icon_news.gif"
	Title_This_Page="Thiết lập sản phẩm bán kèm -> Sửa đổi"
    Call header()
	NewsId	=	GetNumeric(Request.QueryString("NewsId"),0)
	CategoryId	=	GetNumeric(Request.QueryString("catid"),0)
	
%>

<table width="600" border="0" align="center" cellpadding="1" cellspacing="1" class="CTxtContent" style="border:#CCCCCC solid 1">
  <tr>
    <td colspan="2" align="center" valign="bottom" class="CTieuDeNhoNho" >&nbsp;&nbsp;<img src="../../images/icons/Task-List-40x40.gif" width="40" height="40" align="absmiddle"> &nbsp;QUẢN LÝ TIN LIÊN QUAN</td>
  </tr>

  <tr>
    <td colspan="2" background="../../images/line.jpg" style="background-position:bottom;  background-repeat:no-repeat" class="CTieuDeNho">Tin chính</td>
  </tr>
  <tr>
    <td colspan="2" align="left">
	<%
	sqlNews="SELECT * from V_News where NewsId=" & NewsId
       ' Response.Write sqlNews
	Dim rsNews 
	Set rsNews=Server.CreateObject("ADODB.Recordset")
	
	rsNews.open sqlNews,con,3
	if rsNews.eof then
		rsNews.close
		set rsNews=nothing
	else
		idcode=rsNews("Newsid")
		Title=rsNews("Title")
	
		
		PictureId=rsNews("PictureId")
		CategoryId=rsNews("CategoryId")
		PictureAlign=Trim(rsNews("PictureAlign"))
		if UCase(PictureAlign) = "RIGHT" or UCase(PictureAlign)="CENTER" then
			PictureAlign=Trim(rsNews("PictureAlign"))
		else
			PictureAlign	=	"left"
		end if
		rsNews.close
		set rsNews=nothing
	end if	
	set rsPic=server.CreateObject("ADODB.Recordset")
	sql="SELECT * from Picture where PictureId='" & PictureId&"'"
	rsPic.Open sql,con,1
	if not rsPic.eof then
		strNamePic	=	rsPic("SmallPictureFileName")
	end if
	ImagePath	=	NewsImagePath & strNamePic
	%>
	<img src="<%=ImagePath%>" border="0" width="120"  align="left">
	<u>Mã</u>: <%=idcode%><br>
	<u>Tiêu đề</u>: <%=Title%><br>
   </tr>
  <tr>
    <td colspan="2" align="center">
      <input name="OK" type="button" id="OK" value="Cập nhật" onClick="javascript: window.location = 'news_insertsuccess.asp?newsid=<%=NewsId%>&catid=<%=CategoryId%>'">    </td>
  </tr>

  
  <tr>
    <td colspan="2" background="../../images/line.jpg" style="background-position:bottom;  background-repeat:no-repeat" class="CTieuDeNho">Tin liên quan</td>
  </tr>
  <tr>
    <td colspan="2" align="center">
	
<table width="100%" border="0" cellpadding="2" cellspacing="2">
<tr>
<%	
sql	=	"select * from connection where NewsId=" & NewsId		
Set rscon=server.CreateObject("ADODB.Recordset")
rscon.open sql,con,3
do while not rscon.eof
	NewsConnectID	=	rscon("NewsConnectID")

	Set rs=server.CreateObject("ADODB.Recordset")
	sql="SELECT * FROM	V_News where NewsID="&NewsConnectID
	rs.open sql,con,3
	if not rs.eof then
		Title	=rs("Title")
		PictureId	=	rs("PictureId")
		PictureAlign	=	rs("PictureAlign")
		if PictureAlign = "" then
			PictureAlign = "left"
		end if
	%>
	<td valign="top" style="border:#CCCCFF solid 1px;">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" >
		  <tr> 
			<td class="textbody" align="left"><%=Title%>
				  <%
	set rsPic=server.CreateObject("ADODB.Recordset")
	sql="SELECT * from Picture where PictureId=" & PictureId
	rsPic.Open sql,con,1
	if not rsPic.eof then
                      strNamePic	=	rsPic("SmallPictureFileName")
                      %><img src="<%=NewsImagePath&strNamePic%>" border="0" width="150px"  align="middle">	
                <%
		
	end if
	%>
	
    
	
			<div align="right">
                <img src="../../Images/Icons/DeleteRed.png"  width="32" border="0" onClick="javascript: yn = confirm('Bạn có chắn chắn xóa ?'); if(yn) {winpopup('up_attach_news.asp','1&NewsId=<%=NewsId%>&NewsConnectID=<%=NewsConnectID%>&iStatus=Del','100','100')}"/></div>			</td>
		  </tr>
		</table>	</td>			
		
<%
		iCol = iCol + 1
	iF iCol >= 3 then
		iCol	=	0
		Response.Write("</tr><tr>")
	end if
	end if


rscon.movenext
loop
set rscon = nothing		
	%>
</tr>
</table>	</td>
  </tr>
  <tr>
    <td colspan="2" align="center" style="border-bottom:#CCCCCC solid 1">&nbsp;</td>
  </tr>
  <tr>
    <td width="121" align="center"><a href="javascript:winpopup('SearchNews.asp','1&NewsId=<%=NewsId%>','950','400')"><img src="../../images/icons/Default-Icon.png" alt="Thêm mới" border="0" align="absmiddle"> Thêm mới </a></td>
    <td width="472" align="right"></td>
  </tr>
</table>


<%Call Footer()%>

</body>
</html>