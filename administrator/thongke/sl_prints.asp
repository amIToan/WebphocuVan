<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/include/func_tiny.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<%
f_permission = administrator(false,session("user"),"m_report")
if f_permission < 2 then
	response.Redirect("/administrator/info.asp")
end if
%>

<%
	CatID	=	Clng(Trim(Request.QueryString("CatID")))
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=PAGE_TITLE%></title>
<link href="../../css/styles.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table  width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
      <tr>
        <td width="48%"><div align="center"><img src="../../images/logoxseo128.png" width="128"></div></td>
        <td width="53%"align="center" valign="bottom"><em>www.xseo.com</em><br>
            <em>ĐT: <%=soDT%> - Email: info@xseo.com</em></td>
      </tr>
      <tr>
        <td><div align="center"><strong><%=TenGD%></strong></div></td>
        <td width="53%"><div align="center"><em>ĐC: <%=dcVanPhong%>   </em></div></td>
      </tr>
</table>
	<br>
	<br>
<%
	sql = "SELECT  CategoryName, YoungestChildren"
	sql = sql + " FROM NewsCategory"
	sql = sql + " WHERE CategoryID = "&CatID
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open sql,Con,1
	iMumberChild=0
	if not rs.eof then
		iMumberChild	=	Clng(rs("YoungestChildren"))
		strNameCategory	=	rs("CategoryName")	
	end if
	Response.Write("<b>.&nbsp;"&strNameCategory&"</b>")
	Dim arSubCat
	arSubCat = getChilds(CatID)
	if iMumberChild = 0 then
		DisNews(CatID)
	else
		For t = 0 to iMumberChild -1			
			Response.Write("<br><br><b>-&nbsp;"&arSubCat(1,t)&"</b><br>")
			DisNews(arSubCat(0,t))		
		next
	end if
%>

<%
sub DisNews(CategoryTemp)
		sql="SELECT * "
		sql=sql & " FROM V_News n"
		sql=sql & " WHERE  n.CategoryID ='" & CategoryTemp & "'"
		'sql=sql & " and (DATEDIFF(dd, n.CreationDate '" & FromDate & "') <= 0) AND (DATEDIFF(dd, n.CreationDate, '" & ToDate & "') >= 0) "
		sql=sql & " ORDER BY n.NewsId desc"
		Set rsNews = Server.CreateObject("ADODB.Recordset")
		stt=1
		rsNews.open sql,con,1
%>	
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0"  style="border:#000099 solid 1" class="CTxtContent">
	<tr>
		<td width="3%" align="center" bgcolor="#FFFFCC"  style="<%=setStyleBorder(1,1,1,1)%>"><b>STT</b></td>
		<td width="47%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Tên</strong></td>
		<td width="11%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>NXB</strong></td>
		<td width="24%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Tác giả </strong></td>
		<td width="5%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Năm </strong></td>
		<td width="5%" align="center" bgcolor="#FFFFCC" style="<%=setStyleBorder(0,1,1,1)%>"><strong><strong>Gía bìa </strong></strong></td>
		<td width="10%" align="center" bgcolor="#FFFFCC"  style="<%=setStyleBorder(0,1,1,1)%>"><strong>Trạng thái </strong></td>
	    <td width="5%" align="center" bgcolor="#FFFFCC"  style="<%=setStyleBorder(0,1,1,1)%>">Tồn kho </td>
	</tr>
	
<%
		do while not rsNews.eof 
			'Giabia	=	CLng(rsNews("giabia"))
%>
		<tr>
		  <td align="center" style="<%=setStyleBorder(1,1,0,1)%>">
		  <%=stt%>
		  <input name="NewsID<%=stt%>" value="<%=rsNews("NewsID")%>" type="hidden">		  </td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>">
		  <%
		  	strTemp		=	LCase(rsNews("Title"))
			str			=	left(strTemp,1)
			str			=	Ucase(str)
			strTemp		=	str	+ Right(strTemp,len(strTemp)-1)
			Response.Write(strTemp)
		  %>	  </td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>">
		  		  	<%'=rsNews("nxb")%>	  </td>
		  <td style="<%'=setStyleBorder(0,1,0,1)%>" ><%'=rsNews("tacgia")%></td>
		  <td style="<%'=setStyleBorder(0,1,0,1)%>"><%
		  '	strTemp	=	right(rsNews("namxuatban"),1)
			if strTemp <> "" and strTemp <> "0" then
				Response.Write("Quý"&strTemp&"/")
			end if
		  	strTemp	=	Left(rsNews("namxuatban"),4)
			Response.Write(strTemp)
			%></td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>" align="right">
		  <%=Dis_str_money(Giabia)%></td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>"  align="center">
		 
		  <%
		  if rsNews("Het") = 1 then
		  	Response.Write("Hết")
		  else
		  	Response.Write("Còn")
		  end if
		  %>		  </td>
	      <td style="<%=setStyleBorder(0,1,0,1)%>"  align="center"><%=GetNumInventoryGoods(rsNews("NewsID"))%></td>
	  </tr>

<%
			stt	=	stt	+1
			rsNews.movenext
		loop
%>
		<tr>
		  <td align="center" style="<%=setStyleBorder(1,1,0,1)%>">&nbsp;</td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>" align="right">&nbsp;</td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>" align="right">&nbsp;</td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>" align="right">&nbsp;</td>
		  <td style="<%=setStyleBorder(0,1,0,1)%>" align="center">&nbsp;</td>
          <td style="<%=setStyleBorder(0,1,0,1)%>" align="center">&nbsp;</td>
	</tr>
</table>
<%
	end sub
%>
<br>
</body>
</html>
