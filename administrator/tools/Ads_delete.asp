<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
f_permission = administrator(false,session("user"),"m_ads")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
IF IsNumeric(request.Form("Ads_id")) and Clng(request.Form("Ads_id"))<>0 and IsNumeric(request.Form("CatId")) and IsNumeric(request.Form("MultiCat")) THEN
	Ads_id=Clng(request.Form("Ads_id"))
	MultiCat=Clng(request.Form("MultiCat"))
	CatId=Clng(request.Form("CatId"))
	set rs=server.createObject("ADODB.Recordset")
	if MultiCat=1 then
	'Có nhiều chuyên mục, chỉ xóa 1 record ở bảng AdsDistribution
		sql="delete AdsDistribution where Ads_id=" & Ads_id & " and CategoryId=" & CatId
		rs.open sql,con,1
	else
	'Quảng cáo chỉ hiển thị tại một chuyên mục, xóa ở cả hai bảng: Ads và AdsDistribution
		sql="delete AdsDistribution where Ads_id=" & Ads_id & ";delete Ads where Ads_id=" & Ads_id
		rs.open sql,con,1
	end if
	set rs=nothing
	response.Write "<script language=""JavaScript"">" & vbNewline &_
			"<!--" & vbNewline &_
				"window.opener.location.reload();" & vbNewline &_
				"window.opener.focus();" & vbNewline &_
				"window.close();" & vbNewline &_
			"//-->" & vbNewline &_
			"</script>" & vbNewline
	response.End()
ELSE
	if Request.QueryString("param")<>"" and IsNumeric(Request.QueryString("param")) then
		Ads_id=Clng(Request.QueryString("param"))
	else
		response.Redirect("/administrator/default.asp")
		response.End()
	end if
	
	if Request.QueryString("CatId")<>"" and IsNumeric(Request.QueryString("CatId")) then
		CatId=Clng(Request.QueryString("CatId"))
	else
		response.Redirect("/administrator/default.asp")
		response.End()
	end if
	
	sql="SELECT	TOP 2 a.Ads_id, a.Ads_Title " &_
		"FROM	Ads a INNER JOIN " &_
		"			AdsDistribution ad ON a.Ads_id = ad.Ads_id " &_
		"WHERE     (a.Ads_id = " & Ads_id & ")"
		
	set rs=server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
	if not rs.eof then
		Ads_Title=rs("Ads_Title")
		rs.movenext
		if not rs.eof then
			MultiCat=1
		else
			MultiCat=0
		end if
	else
		response.Redirect("/administrator/default.asp")
		response.End()
	end if
	rs.close
	set rs=nothing
END IF
%>
<html>
<head>
    <title><%=PAGE_TITLE%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <link type="text/css" href="../../css/w3.css" rel="stylesheet" />
    <link type="text/css" href="../../css/font-awesome.css" rel="stylesheet" />
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="post" name="fDelete">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr align="center"> 
    <td colspan="2" valign="middle"><font size="2" face="Arial, Helvetica, sans-serif">
    	<strong><%=Ads_Title%></strong>
    </font></td>
  </tr>
  <tr align="center" valign="middle"> 
    <td height="40" colspan="2"><font size="2" face="Arial, Helvetica, sans-serif">
		Bạn chắc chắn muốn xóa Quảng cáo này?
	</font> </td>
  </tr>
  <tr> 
    <td width="50%" height="25" align="center" valign="middle">
        <font size="2" face="Arial, Helvetica, sans-serif">
		<a class="w3-btn w3-red w3-round" href="javascript: window.document.fDelete.submit();">
            <i class="fa fa-trash-o fa-lg" aria-hidden="true"></i> Xóa QCáo
		</a> 
        </font>
    </td>
    <td width="50%" height="25" align="center" valign="middle">
        <font size="2" face="Arial, Helvetica, sans-serif">
            <a class="w3-btn w3-blue w3-round" href="javascript: window.close();"><i class="fa fa-times" aria-hidden="true"></i> Ðóng cửa sổ</a>
        </font>
    </td>
  </tr>
</table>
<input type="hidden" name="Ads_id" value="<%=Ads_id%>">
<input type="hidden" name="MultiCat" value="<%=MultiCat%>">
<input type="hidden" name="CatId" value="<%=CatId%>">
</form>
</body>
</html>
