<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/include/func_tiny.asp" -->
<%
	if Trim(session("user"))="" then
		response.Redirect("/administrator/default.asp")
		response.End()
	end if
%>
<%
	keyword=Trim(Request.Form("keyword"))
	if keyword ="" then
		keyword	=	Request.QueryString("keyword")
	end if
	Keyword=Replace(keyword,"'","''")
	seach_filter=Request.Form("select_filter")
	if seach_filter="" then
		seach_filter=Request.QueryString("seach_filter")
	end if
	NewsConnectID	=	GetNumeric(Request.QueryString("NewsConnectID"),0)
	sNewsID	=	GetNumeric(Request.QueryString("NewsID"),0)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=PAGE_TITLE%></title>
<link href="../../css/styles.css" rel="stylesheet" type="text/css">
<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
<td ></td>
</tr>
<tr>
<td  align="center" class="style18">
<img src="../../images/icons/icon_search.jpg" width="50" height="28" align="absmiddle"><span class="CTieuDeNhoNho"> <%if is_replace=1 then%>TÌM KIẾM SẢN PHẨM THAY THẾ<%else%>TÌM KIẾM<%end iF%></span></td>
</tr>
<tr>
<td >
<FORM action="<%=Request.ServerVariables("SCRIPT_NAME")%>?NewsID=<%=sNewsID%>" method="post" name="fTimKiem" >


<table width="500" border="0" align="center" cellpadding="0" style="border:#CCCCCC solid 1px;" >
<tr>
<td><table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td align="center" width="95%" >
<font size="1" face="verdana"><strong>
Tìm kiếm
</strong></font>
<input name="keyword" type="text"  value="<%=Replace(KeyWord,"""","&quot;")%>" size="25" style="font-family: Verdana; font-size: 8pt; border-style: solid; border-width: 1" onKeyUp="initTyper(this);">---->

<select name="select_filter" id="select_filter">
<option value="1" <% if seach_filter = 1  then Response.Write("selected")%>>Mã sách</option>
<option value="2" selected <% if seach_filter = 2  then Response.Write("selected")%>>Tiêu đề</option>
<option value="3" <% if seach_filter = 3  then Response.Write("selected")%>>Nội dung</option>
</select>
</font> </td>
</tr>
<tr>
<td align="center" ><font size="1" face="verdana">
<input name="search" type="submit" class="style14" style="font-family: Verdana; font-size: 8pt; " value="   Tìm   " >
<input name="SanPhamUser_ID" type="hidden" value="<%=SanPhamUser_ID%>">
<input name="IDSachTB" type="hidden" value="<%=IDSachTB%>">
</font> </td>
</tr>
</table></td>
</tr>
</table>
<%if Keyword<>"" then%>
<%
On Error Resume Next

Dim ArrStringSearch
StringSearch=Trim(Keyword)
Set rsTemp=server.CreateObject("ADODB.Recordset")
sql="SELECT * FROM	V_News "
select case seach_filter
	case 1
		sql= sql+ "Where (idcode like N'%"&StringSearch&"%')"
	case 2
		StringTitle	=	UCASE(StringSearch)
		sql= sql+ "Where (Title like N'%"&StringSearch&"%' or Title like N'%"&StringTitle&"%')"
	case 3
		sql= sql+ "Where (Body like N'%"&StringSearch&"%')"
end select
sql= sql + " ORDER BY NewsID DESC"
  
rsTemp.open sql,con,3
%>
<br />
<diV align="center">Tìm thấy <b><%=rsTemp.recordcount%></b> cuốn sách thoả mãn điều kiện.</div>
<%

if not rsTemp.eof then
%>
<table width="500px" border="0" cellpadding="0" cellspacing="0" align="center"  >
<tr>
<td >Mã</td>
<td>Tiêu đề </td>
<td>&nbsp;</td>
</tr>
<%
Do while not rsTemp.eof 
    NewsID=rsTemp("NewsID")
    idcode = rsTemp("idcode")
    Title	=rsTemp("Title")
%>

<tr>
<td><%=idcode%>
</td>
<td height="18" class="auto-style1">
<%=Title%>					
</td>
<td>
<a href="up_attach_news.asp?NewsConnectID=<%=NewsID%>&NewsID=<%=sNewsID%>&iStatus=add"><img src="../../images/icons/iconAltFormat.gif" border="0" height="40" width="40" align="absmiddle"></a>
</td>
</tr>

<%
rsTemp.movenext
Loop
%>
</table>
<%	
rsTemp.close
set rsTemp=nothing
end if 
END IF
%>

</form>	
</td></tr>
<tr>
<td ></td>
</tr>
</table>
</body>
</html>
