<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script language="javascript">
	function Initialize()
	{
		if (window.opener.fInsert.categoryid.value!="")
		{
			strTmp=window.opener.fInsert.categoryid.value;
			ArrStr=strTmp.split(" ");
			for (j=0;j<ArrStr.length;j++)
			{
				for (i=0;i<document.fChoose.CatId_DependRole.length;i++)
				{
					if (document.fChoose.CatId_DependRole.options[i].value==ArrStr[j])
					{
						document.fChoose.CatId_DependRole.options[i].selected=true;
					}
				}
			}
		}
	}
	function onButtonXoaClick()
	{
		for (i=0;i<document.fChoose.CatId_DependRole.length;i++)
			document.fChoose.CatId_DependRole.options[i].selected=false;
	}
	function onButtonClick()
	{
		strTmp=""
		for (i=0;i<document.fChoose.CatId_DependRole.length;i++)
		{
			if (document.fChoose.CatId_DependRole.options[i].selected)
			{
				strTmp+=" " + document.fChoose.CatId_DependRole.options[i].value;
			}
		}
		//Array.join(delimiter)
		if (strTmp!="")
		{
			window.opener.fInsert.CatId_DependRole.disabled=true;
			window.opener.fInsert.categoryid.value=strTmp;
		}
		else
		{
			window.opener.fInsert.CatId_DependRole.disabled=false;
			window.opener.fInsert.categoryid.value="";
			return;
		}
		self.close();
	}
</script>
</head>
<body onLoad="Initialize();">
<form name="fChoose">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
      <td height="20" align="center" valign="middle"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Chọn 
        nhiều chuyên mục</strong></font></td>
  </tr>
  <tr> 
    <td align="center"><%Call List_Category_MultiChoose_Depend_Role(0, "Danh s&#225;ch chuy&#234;n m&#7909;c","NONE",session("LstRole"),"ap",0)%>
        <br>
        <font size="2" face="Arial, Helvetica, sans-serif"> CTRL + Click để chọn 
        nhiều chuyên mục</font></td>
  </tr>
  <tr> 
    <td align="center">
		<input type="button" name="Button" value=" Chọn " onClick="javascript: onButtonClick();">
        <input type="button" name="Button" value=" Xóa " onClick="javascript:  onButtonXoaClick();">
        <input type="button" name="Button" value=" Đóng " onClick="javascript:  window.close();">
	</td>
  </tr>
</table>
</form>
</body>
</html>
