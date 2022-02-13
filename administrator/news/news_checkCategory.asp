<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Call Authenticate("None")
%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
	Dim Upload 'Su dung AspUpload
	Set Upload = Server.CreateObject("Persits.Upload")

	Upload.SetMaxSize 1000000, True 'Dat kich co upload la` 1MB
	Upload.codepage=65001
	Upload.Save
		LstCat=Trim(Replace(Upload.Form("categoryid"),"'","''"))
		if LstCat<>"" then
			Arr=split(" " & LstCat)
		else
			LstCat=Trim(Replace(Upload.Form("CatId_DependRole"),"'","''"))
			Arr=split(" " & LstCat)
		end if
	set Upload=nothing
%>
<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="fCheck" method="post">
<table width="100%" align="center" border="0"><tr align="center" valign="middle"><td height="25">
<font size="2" face="Arial, Helvetica, sans-serif"><strong>
	Trạng thái tin cho từng chuyên mục
</strong></font>
 </td>
</tr>
<tr><td>
  <table width="100%" border="0" cellpadding="0" cellspacing="1" bordercolor="#000000" bgcolor="#000000" align="center">
  <%for i=1 to UBound(Arr)%>
    <tr <%if i mod 2=0 then%>bgcolor="#FFFFFF"<%else%>bgcolor="#E6E8E9"<%end if%>> 
      <td><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;<%=GetListParentCatNameOfCatId(Arr(i))%></font></td>
	  <%Select case GetRoleOfCat_FromListRole(Arr(i),session("LstRole"))
	  	case "ed"
			Response.Write "<td><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_ 
		    "	<input type=""radio"" name=""Cat" & Arr(i) & """ value=""edma"">Đánh dấu" & vbNewline &_ 
			"</font></td>" & vbNewline &_ 
      		"<td colspan=""2""><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_ 
	        "	<input type=""radio"" name=""Cat" & Arr(i) & """ value=""edse"">Gửi lên Hiệu đính" & vbNewline &_ 
			"</font></td>" & vbNewline
		case "se"
			Response.Write "<td><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_ 
		    "	<input type=""radio"" name=""Cat" & Arr(i) & """ value=""sema"">Đánh dấu" & vbNewline &_ 
			"</font></td>" & vbNewline &_ 
      		"<td colspan=""2""><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_ 
	        "	<input type=""radio"" name=""Cat" & Arr(i) & """ value=""seap"">Gửi lên Phụ trách" & vbNewline &_ 
			"</font></td>" & vbNewline
		case "ap"
			Response.Write "<td><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_ 
		    "	<input type=""radio"" name=""Cat" & Arr(i) & """ value=""apma"">Đánh dấu" & vbNewline &_ 
			"</font></td>" & vbNewline &_ 
      		"<td><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_ 
	        "	<input type=""radio"" name=""Cat" & Arr(i) & """ value=""apap"">Gửi lên mạng" & vbNewline &_ 
			"</font></td>" & vbNewline &_
			"<td><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_
	        "	<input type=""radio"" name=""Cat" & Arr(i) & """ value=""apad"">Gửi lên Tổng phụ trách" & vbNewline &_
			"</font></td>" & vbNewline
		case "ad"
			Response.Write "<td><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_ 
		    "	<input type=""radio"" name=""Cat" & Arr(i) & """ value=""adma"">Đánh dấu" & vbNewline &_ 
			"</font></td>" & vbNewline &_ 
      		"<td colspan=""2""><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_ 
	        "	<input type=""radio"" name=""Cat" & Arr(i) & """ value=""adad"">Gửi lên mạng" & vbNewline &_ 
			"</font></td>" & vbNewline
	  End Select%>
    </tr>
	<%Next%>
  </table>
 </td>
</tr>
<tr>
    <td align="center" valign="bottom" height="30"><input type="button" name="Button" value=" Chọn " onClick="javascript: CheckSelected();">
      <input type="button" name="Submit2" value=" Đóng " onClick="javascript: window.close();"></td>
</tr></table>
</form>
<script language="JavaScript">
	function CheckSelected()
	{
		var i;
		var formElements = document.fCheck.elements;
		var varname="";
		var boo=false;
		var strcategoryid="";
		var strnews_category_status="";

		for (i=0; i<formElements.length; i++) 
		{
			if (formElements[i].type == "radio")
			{
			  	if ((i==0) || (varname!=formElements[i].name))
			  	{
					varname=formElements[i].name;
					
					for (j=0;j<eval('document.fCheck.' + varname + '.length');j++)
					{
						if (eval('document.fCheck.' + varname + '[' + j + '].checked'))
						{
							//alert(varname + "_" + j + "=" + eval('document.fCheck.' + varname + '[' + j + '].value'));
							strnews_category_status+=" " + eval('document.fCheck.' + varname + '[' + j + '].value');
							//Arr=varname.split("Cat");
							//alert (Arr.length + ",0:" + Arr[0] + ",2:" + Arr[1]);
							strcategoryid+=" " + varname.split("Cat")[1];
							
							boo=false;
							break;
						}
						else
							boo=true;
					}//for (j=0;j<eval('document.fCheck.' + varname + '.length');j++)
					if (boo) 
					{  
						alert("Bạn phải chọn trạng thái tin cho từng chuyên mục."); 
						return;
					}
				}//if ((i==0) || (varname!=formElements[i].name))
			}//if (formElements[i].type == "radio")
		}//for (i=0; i<formElements.length; i++) 
		window.opener.fInsert.news_category_status.value=strnews_category_status;
		window.opener.fInsert.categoryid.value=strcategoryid;
		window.opener.SendNews();
		//window.opener.document.fInsert.news_category_status.value=strnews_category_status;
		//window.opener.document.fInsert.categoryid.value=strcategoryid;
		//alert ("he he");
		//window.opener.document.SendNews();
		//window.opener.document.fInsert.action="news_insert.asp";
		//alert (window.opener.document.fInsert.action);
		//window.opener.document.fInsert.submit();
		self.close();
	}
</script> 
</body>
</html>
