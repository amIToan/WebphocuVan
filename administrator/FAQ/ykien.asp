<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
f_permission = administrator(false,session("user"),"m_faq")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
	lang=Request.QueryString("param")
	lang=replace(lang,"'","''")
%>
	
<%
	idYKien = Clng(Request.QueryString("id"))
	action  = Request.QueryString("action")
Select case action 
	case "edit"
		f_permission = administrator(false,session("user"),"m_faq")
		if f_permission <= 2 then
			response.Redirect("/administrator/info.asp")
		end if	
		sqlyk = "Select * From Y_KIEN where id='"& idYKien &"'"
		Set rsYKien=Server.CreateObject("ADODB.Recordset")
		rsYKien.open sqlyk,con,1
		if not rsYKien.eof then
			Email=rsYKien("email")
			hovaten=Trim(rsYKien("hovaten"))
			tieude=	Trim(rsYKien("tieude"))
			noidung=Trim(rsYKien("noidung"))
			faq=rsYKien("faq")
			isshow=Clng(rsYKien("show"))
			ngaytao=rsYKien("ngaytao")
			Traloi	=	rsYKien("Traloi")
		end if
		set rsYKien = nothing
'		call UserOperation(session("user"),hour(now)&":"&Minute(now)&"phút: sửa ý kiến mã: "&idYKien)
	case "del"
		f_permission = administrator(false,session("user"),"m_faq")
		if f_permission <= 3 then
			response.Redirect("/administrator/info.asp")
		end if		
		sqlyk = "Delete Y_KIEN where id='"& idYKien &"'"
		Set rsYKien=Server.CreateObject("ADODB.Recordset")
		rsYKien.open sqlyk,con,1
		set rsYKien = nothing
		call UserOperation(session("user"),hour(now)&":"&Minute(now)&"phút: xóa ý kiến mã: "&idYKien)
		Response.Write	"<script language=""JavaScript"">" & vbNewline &_
		"	<!--" & vbNewline &_
		"		window.opener.location.reload();" & vbNewline &_
		"		window.close();" & vbNewline &_
		"	//-->" & vbNewline &_
		"</script>" & vbNewline

	case "update"
		f_permission = administrator(false,session("user"),"m_faq")
		if f_permission <= 2 then
			response.Redirect("/administrator/info.asp")
		end if	
		Email=Trim(Request.Form("txtEmail"))
		hovaten=Trim(Request.Form("txtName"))
		tieude=	Trim(Request.Form("txtTitle"))
		noidung=Trim(Request.Form("txtAns"))
        tieude=replace(tieude,"'","''")
        noidung=replace(noidung,"'","''")

		isshow=	Clng(Request.Form("isShow"))
		if isshow <> 1 then
			isshow = 0
		else
			isshow = 1
		end if 
		Traloi	=	Trim(Request.Form("txtQuestion"))
	
		sqlyk = "Update Y_KIEN set hovaten = N'"& hovaten &"', email='"& Email &"', tieude =N'"& tieude &"', noidung=N'"& noidung &"', show = '"& isshow &"', Traloi = N'"& Traloi &"' where id='"& idYKien &"'"
		Set rsYKien=Server.CreateObject("ADODB.Recordset")
		rsYKien.open sqlyk,con,1
		
		set rsYKien = nothing
		call UserOperation(session("user"),hour(now)&":"&Minute(now)&"phút: sửa ý kiến mã: "&idYKien)
		Response.Write	"<script language=""JavaScript"">" & vbNewline &_
		"	<!--" & vbNewline &_
		"		window.opener.location.reload();" & vbNewline &_
		"		window.close();" & vbNewline &_
		"	//-->" & vbNewline &_
		"</script>" & vbNewline
end select
%>
<html>
<head>
    <title><%=PAGE_TITLE%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <link href="../../css/styles.css" rel="stylesheet" type="text/css">
    <script src="../../ckeditor/ckeditor.js"></script>
    <script src="../../ckfinder/ckfinder.js"></script>
    <link type="text/css" href="../../css/w3.css" rel="stylesheet" />
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="ykien.asp?id=<%=idYKien%>&action=update">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="w3-table w3-table-all w3-hoverable">
  <tr>
    <td colspan="2" align="center"class="CTitleClass_AI"><br>
    SỬA Ý KIẾN KHÁCH HÀNG<br></td>
    </tr>
  <tr>
    <td align="right">Họ và tên: </td>
    <td><input name="txtName" type="text" id="txtName" value="<%=hovaten%>" size="25" class="w3-input w3-border w3-round"/></td>
  </tr>
  <tr>
    <td align="right">Email: </td>
    <td><input name="txtEmail" type="text" id="txtEmail" value="<%=Email%>" size="25" class="w3-input w3-border w3-round"/></td>
  </tr>
  <tr>
    <td align="right">Tiêu đề: </td>
    <td><input name="txtTitle" type="text" id="txtTitle" value="<%=tieude%>" size="25" class="w3-input w3-border w3-round"/></td>
  </tr>
  <tr>
    <td width="24%" align="right">Ý kiến  : </td>
    <td width="76%">
      <textarea name="txtAns" cols="50" rows="10" id="txtAns" class="w3-input w3-border w3-round"><%=noidung%></textarea>   </td>
  </tr>
  <tr>
    <td width="24%" align="right">Trả lời  : </td>
    <td width="76%">
      <textarea name="txtQuestion" cols="50" rows="10" id="txtQuestion" class="w3-input w3-border w3-round"><%=Traloi%></textarea></td>
  </tr>
  <tr>
    <td align="right">Cho phép hiển thị:</td>
    <td><input name="isShow" type="checkbox" id="isShow" class="w3-check" value="1" <% if isshow <> 0 then %> checked="checked" <%end if%>></td>
  </tr>
  <tr>
    <td colspan="2" align="right"><div align="center">
      <input class="w3-btn w3-red w3-round" type="submit" name="Submit" value="    sửa    " onClick="CheckMe();">
    </div></td>
    </tr>
</table>
</form>
<script type="text/javascript">
    CKEDITOR.replace('txtAns');
    CKEDITOR.replace('txtQuestion');
</script>
</body>
</html>
<script language="javascript">
	if (document.form1.txtName.value==0)
	{
		alert("Mời bạn nhập họ và tên người gửi!");
		document.form1.txtName.focus();
		return false;
	}
	if (document.form1.txtEmail.value=='')
	{
		alert("Mời bạn nhập email");
		document.form1.txtEmail.focus();
		return false;
	}
	if (document.form1.txtTilte.value==0)
	{
		alert("Mời bạn nhập tiêu đề!");
		document.form1.txtTilte.focus();
		return false;
	}
	if (document.form1.txtAns.value==0)
	{
		alert("Mời bạn nhập nội dung câu hỏi!");
		document.form1.txtAns.focus();
		return false;
	}
</script>