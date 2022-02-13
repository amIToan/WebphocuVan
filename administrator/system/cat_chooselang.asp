<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Call Authenticate("Admin")%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
	lang=Request.QueryString("param")
	lang=replace(lang,"'","''")
%>
<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link type="text/css" href="../../bootstrap/css/w3.css" rel="stylesheet" />
<script language="JavaScript">
	<!--
		function chooselang(thisvalue)
		{
			window.opener.location="cat_list.asp?lang=" + thisvalue;
			window.close();
		}
	//-->
</script>
</head>
<body>
<form name="fChooseLang" method="post">
  <table width="96%" border="0" cellspacing="2" cellpadding="2" align="center">
    <tr align="center" valign="middle"> 
      <td height="30" colspan="2"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Chọn 
        ngôn ngữ</strong></font></td>
    </tr>
    <tr> 
      <td><font size="2" face="Arial, Helvetica, sans-serif">Ngôn ngữ:</font></td>
      <td> <%Call List_Language(lang)%> </td>
    </tr>
    <tr align="center"> 
	<td colspan="2" align="center"> 
		<input class="w3-btn w3-red" type="Button" name="Button" value="  Chọn  " onClick="javascript: chooselang(document.fChooseLang.languageid.value)">
        <input class="w3-btn w3-blue" type="button" name="Submit2" value="  Đóng  " onClick="javascript: window.close();">
      </td>
    </tr>
  </table>
</form>
</body>
</html>
