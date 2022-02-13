<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Call PhanQuyen("QLyHeThong")%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->

<%
	lang=Request.QueryString("param")
	lang=replace(lang,"'","''")
	
	CatId=Request.QueryString("Catid")
	if not isNumeric(catid) then
		response.End()
	end if
	
	if request.Form("confirm")="Yes" and catId<>0 then
'<ngày 14/6/2007>
		set rss=server.CreateObject("ADODB.Recordset")
		sql="Select * From NewsDistribution where CategoryID="& CatId
		rss.open sql,con,1
		do while not rss.eof
			newid=rss("NewsID")
			sql="Delete news where NewsID="& newid
			set rrs=server.CreateObject("ADODB.Recordset")
			rrs.open sql,con,1
		rss.movenext
		loop
'</ngày>		
		Dim rs
		set rs=server.CreateObject("ADODB.Recordset")
		sql="delete Newscategory where categoryId=" & CatId
		rs.open sql,con,1
		set rs=nothing
		set rrs=nothing
		set rss=nothing
		
		Call Update_PrentCategoryId(lang)
		Call Update_YoungestChildren(lang)
		
		Response.Write	"<script language=""JavaScript"">" & vbNewline &_
			"	<!--" & vbNewline &_
			"		window.opener.location.reload();" & vbNewline &_
			"		window.close();" & vbNewline &_
			"	//-->" & vbNewline &_
			"</script>" & vbNewline
		response.End()
	end if
	
	Dim cm
	set cm = CreateObject("ADODB.Command")
    cm.ActiveConnection = strConnString

	cm.commandtype=4 'adstoredProc
	cm.CommandText = "CheckCatFree"
	cm.Parameters.Append cm.CreateParameter("CatId", 3, 1, 4, CatId )
	'Set objparameter=objcommand.CreateParameter (name,type,direction,size,value)
	cm.Parameters.Append cm.CreateParameter("CatNum", 3, 2, 4)
	cm.execute()
	
	CatNum = Cint(cm("CatNum"))
	' ngay 14/6/2007 tạm thời cho phép xóa all
	CatNum=0
	set cm=nothing
%>
<html>
<head>
<title>Chuyen muc moi</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link type="text/css" href="../../css/w3.css" rel="stylesheet" />
<link type="text/css" href="../../css/font-awesome.css" rel="stylesheet" />
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%if CatNum=0 then%>
<form name="fDelete" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>?<%=Request.ServerVariables("QUERY_STRING")%>">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr align="center" valign="middle"> 
    <td height="30"><font size="3" face="Arial, Helvetica, sans-serif">
		<strong><br>Chắc chắn xóa chuyên mục?<br></strong>
		&#8226;&nbsp;<%=GetNameOfCategory(Catid)%>
	</font></td>
  </tr>
  <tr align="center" valign="middle">
    <td height="30"><br><font size="2" face="Arial, Helvetica, sans-serif">
		<a class="w3-btn w3-red w3-round" href="javascript: document.fDelete.submit();"><i class="fa fa-trash-o fa-lg" aria-hidden="true"></i> Chắc chắn</a>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<a class="w3-btn w3-blue w3-round" href="javascript: window.close();"><i class="fa fa-times" aria-hidden="true"></i> Đóng cửa sổ</a>
	</font></td>
  </tr>
</table>
<input type="hidden" name="confirm" value="Yes">
</form>
<%Else%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr align="center" valign="middle"> 
    <td height="30"><br><font size="2" face="Arial, Helvetica, sans-serif"><strong> 
      Không thể xóa được chuyên mục<br><br>
      </strong></font> <font size="2" face="Arial, Helvetica, sans-serif">
	  	Bạn phải xóa hết tin, ảnh, tài khoản đăng nhập<br>liên quan đến chuyên mục này. 
      </font> </td>
  </tr>
  <tr align="center" valign="middle">
    <td height="30"><br><font size="2" face="Arial, Helvetica, sans-serif"><a class="w3-btn w3-blue" href="javascript: window.close();">Đóng cửa sổ</a></font></td>
  </tr>
</table>
<%End if%>
</body>
</html>
