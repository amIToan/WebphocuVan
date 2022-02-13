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
		old_categoryid=Trim(Replace(Upload.Form("old_categoryid"),"'","''"))
		old_news_category_status=Trim(Replace(Upload.Form("old_news_category_status"),"'","''"))
		categoryid=Trim(Replace(Upload.Form("categoryid"),"'","''"))
		All_categoryid=Trim(Replace(Upload.Form("All_categoryid"),"'","''"))
		NewsId=Clng(Upload.Form("NewsId"))
		if categoryid<>"" then
			Arr=split(" " & categoryid)
		else
			categoryid=Trim(Replace(Upload.Form("CatId_DependRole"),"'","''"))
			Arr=split(" " & categoryid)
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
  <%FOR i=1 TO UBound(Arr)%>
    <tr <%if i mod 2=0 then%>bgcolor="#FFFFFF"<%else%>bgcolor="#E6E8E9"<%end if%>> 
      <td><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;<%=GetListParentCatNameOfCatId(Arr(i))%></font></td>
	  <%
	  if Instr(" " & old_categoryid & " "," " & Arr(i) & " ")>0 then
	  'Đã tồn tại trong List cũ.
	   	StatusCheck=GetStatusNameOfNewsAndCategoryId(Arr(i),old_news_category_status)
	   	Select case GetRoleOfCat_FromListRole(Arr(i),session("LstRole"))
	   		case "ed"
	  			Call WriteRadioNewsStatus(NewsId,Arr(i),"Editor",StatusCheck)
	  		case "se"
	  			Call WriteRadioNewsStatus(NewsId,Arr(i),"GroupSenior",StatusCheck)
	  		case "ap"
	  			Call WriteRadioNewsStatus(NewsId,Arr(i),"Approver",StatusCheck)
	  		case "ad"
	  			Call WriteRadioNewsStatus(NewsId,Arr(i),"Administrator",StatusCheck)
	  	end select
	  else
	  'Không nằm trong Danh sách cũ.
	  	if CheckNewsAndCategoryIdIsExist(NewsId,Arr(i))>0 then
	  	'Đã tồn tại NewsId va` CatId trong bảng NewsDistribution
	  		Response.Write "<td colspan=""4""><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_ 
	  		"Tin đã được người khác đưa vào chuyên mục này." & vbNewline &_
	  		"</font></td>" & vbNewline
	  	else
	  	'Chưa tồn tại NewsId và CatId trong bảng NewsDistribution
	  		Select case GetRoleOfCat_FromListRole(Arr(i),session("LstRole"))
	  		case "ed"
				Response.Write "<td><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_ 
		    	"	<input type=""radio"" name=""Cat" & Arr(i) & """ value=""edma"">Đánh dấu" & vbNewline &_ 
				"</font></td>" & vbNewline &_ 
    	  		"<td colspan=""3""><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_ 
	    	    "	<input type=""radio"" name=""Cat" & Arr(i) & """ value=""edse"">Gửi lên Hiệu đính" & vbNewline &_ 
				"</font></td>" & vbNewline
			case "se"
				Response.Write "<td><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_ 
			    "	<input type=""radio"" name=""Cat" & Arr(i) & """ value=""sema"">Đánh dấu" & vbNewline &_ 
				"</font></td>" & vbNewline &_ 
	      		"<td colspan=""3""><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_ 
		        "	<input type=""radio"" name=""Cat" & Arr(i) & """ value=""seap"">Gửi lên Phụ trách" & vbNewline &_ 
				"</font></td>" & vbNewline
			case "ap"
				Response.Write "<td><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_ 
			    "	<input type=""radio"" name=""Cat" & Arr(i) & """ value=""apma"">Đánh dấu" & vbNewline &_ 
				"</font></td>" & vbNewline &_ 
	      		"<td><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_ 
		        "	<input type=""radio"" name=""Cat" & Arr(i) & """ value=""apap"">Gửi lên mạng" & vbNewline &_ 
				"</font></td>" & vbNewline &_
				"<td colspan=""2""><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_
		        "	<input type=""radio"" name=""Cat" & Arr(i) & """ value=""apad"">Gửi lên Tổng phụ trách" & vbNewline &_
				"</font></td>" & vbNewline
			case "ad"
				Response.Write "<td><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_ 
			    "	<input type=""radio"" name=""Cat" & Arr(i) & """ value=""adma"">Đánh dấu" & vbNewline &_ 
				"</font></td>" & vbNewline &_ 
	      		"<td colspan=""3""><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_ 
		        "	<input type=""radio"" name=""Cat" & Arr(i) & """ value=""adad"">Gửi lên mạng" & vbNewline &_ 
				"</font></td>" & vbNewline
	  		End Select
	  	end if
	  End if%>
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
<%Function CheckNewsAndCategoryIdIsExist(NewsId,CatId)
	sql="select count(NewsId) as dem from NewsDistribution where newsid=" & newsId & " And CategoryId=" & CatId
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
		CheckNewsAndCategoryIdIsExist=CLng(rs("dem"))
	rs.close
	set rs=nothing
End Function%>
<%Function GetStatusNameOfNewsAndCategoryId(CatId,LstStatus)
	Dim sCatId,sLstStatus
	sCatId=" " & CatId
	sLstStatus=" " & sLstStatus & " "
	if Instr(sLstStatus,sCatId & "edma ")>0 then 'Đánh dấu
		StatusNameOfNewsAndCategoryId="edma"
	elseif Instr(sLstStatus,sCatId & "edse ")>0 then 'Gửi lên Hiệu đính
		StatusNameOfNewsAndCategoryId="edse"
	elseif Instr(sLstStatus,sCatId & "seed ")>0 then 'Hiệu đính yêu cầu sửa
		StatusNameOfNewsAndCategoryId="seed"
	elseif Instr(sLstStatus,sCatId & "sema ")>0 then 'Hiệu đính đánh dấu
		StatusNameOfNewsAndCategoryId="sema"
	elseif Instr(sLstStatus,sCatId & "seap ")>0 then 'Hiệu đính gửi lên phụ trách
		StatusNameOfNewsAndCategoryId="seap"
	elseif Instr(sLstStatus,sCatId & "apse ")>0 then 'Phụ trách gửi lại Hiệu đính
		StatusNameOfNewsAndCategoryId="apse"
	elseif Instr(sLstStatus,sCatId & "apma ")>0 then 'Phụ trách đánh dấu
		StatusNameOfNewsAndCategoryId="apma"
	elseif Instr(sLstStatus,sCatId & "apap ")>0 then 'Phụ trách cho Online
		StatusNameOfNewsAndCategoryId="apap"
	elseif Instr(sLstStatus,sCatId & "apad ")>0 then 'Phụ trách gửi lên Tổng phụ trách
		StatusNameOfNewsAndCategoryId="apad"
	elseif Instr(sLstStatus,sCatId & "adap ")>0 then 'Tổng phụ trách gửi lại Phụ trách
		StatusNameOfNewsAndCategoryId="adap"
	elseif Instr(sLstStatus,sCatId & "adma ")>0 then 'Tổng phụ trách đánh dấu
		StatusNameOfNewsAndCategoryId="adma"
	elseif Instr(sLstStatus,sCatId & "adad ")>0 then 'Tổng phụ trách cho Online
		StatusNameOfNewsAndCategoryId="adad"
	end if
End Function%>

<%Sub WriteRadioNewsStatus(NewsId,CatId,UserRole,StatusCheck)
	'UserRole: Editor, GroupSenior, Approver, Administrator
	'StatusChecke: Cho các giá trị: apma, apde, apap, apad
	'Trường hợp UserRole=Approver
SELECT CASE UserRole
CASE "Editor"
	Response.Write "<td><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_ 
	"	<input type=""radio"" name=""Cat" & CatId & """ value=""edma"""
	if StatusCheck="edma" then
		Response.write " checked"
	end if
	Response.write "<td colspan=""3""><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_
	"	<input type=""radio"" name=""Cat" & CatId & """ value=""edse"""
	if StatusCheck="edse" then
		Response.write " checked"
	end if
	Response.write ">Gửi lên Hiệu đính viên" & vbNewline &_
	"</font></td>" & vbNewline
CASE "GroupSenior"
	if not IsNULL(CheckUserProcess(UserRole,CatId,NewsId)) then
		Response.Write "<td><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_ 
		"	<input type=""radio"" name=""Cat" & CatId & """ value=""seed"""
		if StatusCheck="seed" then
			Response.write " checked"
		end if
		Response.write ">Gửi lại Biên tập viên" & vbNewline &_ 
		"</font></td>" & vbNewline
	else
		Response.Write "<td><font size=""2"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & vbNewline
	end if
	Response.Write "<td><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_ 
	"	<input type=""radio"" name=""Cat" & CatId & """ value=""sema"""
	if StatusCheck="sema" then
		Response.write " checked"
	end if
	Response.write ">Đánh dấu " & vbNewline &_ 
		"</font></td>" & vbNewline
	Response.write "<td colspan=""2""><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_
	"	<input type=""radio"" name=""Cat" & CatId & """ value=""seap"""
	if StatusCheck="seap" then
		Response.write " checked"
	end if
	Response.write ">Gửi lên Phụ trách" & vbNewline &_
	"</font></td>" & vbNewline
CASE "Approver"
	if not IsNULL(CheckUserProcess(UserRole,CatId,NewsId)) then
		Response.Write "<td><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_ 
		"	<input type=""radio"" name=""Cat" & CatId & """ value=""apse"""
		if StatusCheck="apse" then
			Response.write " checked"
		end if
		Response.write ">Gửi lại Hiệu đính" & vbNewline &_ 
		"</font></td>" & vbNewline
	else
		Response.Write "<td><font size=""2"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & vbNewline
	end if
	Response.Write "<td><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_ 
	"	<input type=""radio"" name=""Cat" & CatId & """ value=""apma"""
	if StatusCheck="apma" then
		Response.write " checked"
	end if
	Response.write ">Đánh dấu" & vbNewline &_ 
	"</font></td>" & vbNewline &_ 
    "<td><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_ 
	"	<input type=""radio"" name=""Cat" & CatId & """ value=""apap"""
	if StatusCheck="apap" then
		Response.write " checked"
	end if
	Response.write ">Gửi lên mạng" & vbNewline &_ 
	"</font></td>" & vbNewline &_
	"<td><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_
	"	<input type=""radio"" name=""Cat" & CatId & """ value=""apad"""
	if StatusCheck="apad" then
		Response.write " checked"
	end if
	Response.write ">Gửi lên Tổng phụ trách" & vbNewline &_
	"</font></td>" & vbNewline
CASE "Administrator"
	if not IsNULL(CheckUserProcess(UserRole,CatId,NewsId)) then
		Response.Write "<td><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_ 
		"	<input type=""radio"" name=""Cat" & CatId & """ value=""adap"""
		if StatusCheck="adap" then
			Response.write " checked"
		end if
		Response.write ">Gửi lại Phụ trách" & vbNewline &_ 
		"</font></td>" & vbNewline
	else
		Response.Write "<td><font size=""2"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & vbNewline
	end if
	Response.Write "<td><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_ 
	"	<input type=""radio"" name=""Cat" & CatId & """ value=""adma"""
	if StatusCheck="adma" then
		Response.write " checked"
	end if
	Response.write "<td colspan=""2""><font size=""2"" face=""Arial, Helvetica, sans-serif"">" & vbNewline &_
	"	<input type=""radio"" name=""Cat" & CatId & """ value=""seap"""
	if StatusCheck="seap" then
		Response.write " checked"
	end if
	Response.write ">Gửi lên Phụ trách" & vbNewline &_
	"</font></td>" & vbNewline
END SELECT
End Sub%>