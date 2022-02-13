<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Call PhanQuyen("QLyHeThong")%>
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
<script language="JavaScript">
	<!--
		function chooselang(thisvalue1,thisvalue2,thischeck)
		{
			if (thischeck.checked==true)
				window.opener.location="cat_list.asp?lang=" + thisvalue1+"&CategoryLoai=none";
			else
				window.opener.location="cat_list.asp?lang=" + thisvalue1+"&CategoryLoai="+thisvalue2;
			window.close();
		}
	//-->
</script>
<link href="../../css/styles.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="fChooseLang" method="post">
  <table width="96%" border="0" align="center" cellpadding="2" cellspacing="2" class="CTxtContent">
    <tr align="center" valign="middle"> 
      <td height="30" colspan="2"><font size="3" face="Arial, Helvetica, sans-serif"><strong>CHỌN LỌC</strong></font></td>
    </tr>
    <tr>
      <td>Chọn loại: 
	  </td>
	  <td>
	  <%Call ListLoaiOfCategory(0)%>
	  
      Tất cả
      <input type="checkbox" name="AllCheck" value="1" id="AllCheck"></td>
    </tr>
    <tr>
      <td width="24%">&nbsp;</td>
      <td width="76%">&nbsp;</td>
    </tr>
    <tr> 
      <td><font size="2" face="Arial, Helvetica, sans-serif">Ngôn ngữ:</font></td>
      <td> <%Call List_Language(lang)%> </td>
    </tr>
    <tr align="center"> 
	<td colspan="2" align="center"> 
		<input type="Button" name="Button" value="  Chọn  " onClick="javascript: chooselang(document.fChooseLang.languageid.value,document.fChooseLang.CategoryLoai.value,document.fChooseLang.AllCheck)">
        <input type="button" name="Submit2" value="  Đóng  " onClick="javascript: window.close();">      </td>
    </tr>
    <tr align="center">
      <td colspan="2" align="center">&nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>
