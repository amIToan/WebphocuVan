<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/include/func_GetVariables.asp" -->
<%
f_permission = administrator(false,session("user"),"m_customer")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
	<%
	cmnd = Request.QueryString("param")
	sqlEdit	=	"select * from Account where CMND='"& cmnd &"'"
	set rsEdit	=	Server.CreateObject("ADODB.recordset")
	rsEdit.open sqlEdit,con,1
	if not rsEdit.eof then
		ngaycap 	= 	rsEdit("ngaycap")
		dayCap		=	Day(ngaycap) 
		morCap		=	Month(ngaycap) 
		yearCap		=	Year(ngaycap) 
		
		noicap		=	rsEdit("noicap")
		name		=	rsEdit("name")

		ngaysinh	=	rsEdit("ngaysinh")
		f_NgaySinh	=	Day(ngaysinh)
		f_ThangSinh	=	month(ngaysinh)
		f_NamSinh	=	year(ngaysinh)		

		nguyenquan	=	rsEdit("nguyenquan")
		ProvinceID		=	Clng(rsEdit("ProvinceID"))				
		DistrictID		=	Clng(rsEdit("DistrictID"))						
		diachi		=	rsEdit("diachi")
		Email		=	rsEdit("Email")
		Tel			=	rsEdit("Tell")
		mobile		=	rsEdit("mobile")
		pass		=	Trim(rsEdit("password"))
	end if

	if isnumeric(Request.QueryString("reTinh")) = true or Request.QueryString("reTinh") <> "" then
		reTinh	= Request.QueryString("reTinh")
	else
		reTinh= 0
	end if	
		
	if reTinh = 1 then
		cmnd		=	Request.Form("f_cmnd")
		noicap		=	Request.Form("f_noicap")
		dayCap		=	Clng(Request.Form("dayCap"))
		morCap		=	Clng(Request.Form("morCap"))
		yearCap		=	Clng(Request.Form("yearCap"))
		name		=	Request.Form("f_name")
		f_NgaySinh	=	Clng(Request.Form("f_NgaySinh"))
		f_ThangSinh	=	Clng(Request.Form("f_ThangSinh"))
		f_NamSinh	=	Clng(Request.Form("f_NamSinh"))
		nguyenquan	=	Request.Form("f_quequan")
		diachi		=	Request.Form("f_diachi")
		Email		=	Request.Form("f_mail")
		Tel			=	Request.Form("f_tell")
		mobile		=	Request.Form("f_mobile")
		pass		=	Trim(Request.Form("f_pass"))
		ProvinceID 		= 	Clng(Request.Form("selTinh1"))
		DistrictID 		= 	Clng(Request.Form("selHuyen1"))
	end if

%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Hello</title>
<link href="../../css/styles.css" rel="stylesheet" type="text/css">
</head>
<script language="JavaScript" src="include/vietuni.js"></script>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<form name="form_reg" method="post" action="regclose.asp">
<table width="100%" border="0" cellpadding="2" cellspacing="2">
	<tr>
	  <td colspan="3"><div class="style14"> 
	    <div align="center"><strong>
	      SỬA THÔNG TIN </strong></div>
	  </div></td>
	</tr>

	<tr>
	  <td width="39%" height="25"><div align="right"><span class="style7">Chứng minh nhân dân : </span></div></td>
	  <td width="38%"><input name="f_cmnd" type="text" id="f_cmnd" value="<%=cmnd%>" size="35"  class="CTextBoxUnder" readonly="true">	   </td>
	  <td width="23%"><span class="style8">(*)</span></td>
	</tr>
	<tr>
	  <td height="29"><div align="right" class="style7">
		  <div align="right">Ngày cấp : </div>
	  </div></td>
	  <td><% 
		call ex_Day_morth_year(dayCap,morCap,yearCap,"dayCap","morCap","yearCap") 
	%>                              </td>
	  <td>&nbsp;</td>
	</tr>
	<tr>
	  <td height="26"><div align="right" class="style7">
		  <div align="right">Nơi cấp : </div>
	  </div></td>
	  <td><input name="f_noicap" type="text" id="f_noicap" value="<%=noicap%>" size="35" class="CTextBoxUnder"></td>
	  <td>&nbsp;</td>
	</tr>
	<tr>
	  <td height="27"><div align="right"><span class="style7">Họ và tên : </span></div></td>
	  <td><input name="f_name" type="text" id="f_name" value="<%=name%>" size="35" class="CTextBoxUnder"></td>
	  <td><span class="style8">(*)</span></td>
	</tr>
	<tr>
	  <td height="24" class="style7"><div align="right">Ngày tháng năm sinh : </div></td>
	  <td><%
	call ex_Day_morth_year(f_NgaySinh,f_ThangSinh,f_NamSinh,"f_NgaySinh","f_ThangSinh","f_NamSinh") 
%>                              </td>
	  <td>&nbsp;</td>
	</tr>
	<tr>
	  <td height="27"><div align="right"><span class="style7">Quê quán: </span></div></td>
	  <td><input name="f_quequan" id="f_quequan" size="35" class="CTextBoxUnder" value="<%=nguyenquan%>">                              </td>
	</tr>
	
	<tr>
	  <td height="27"><div align="right">Tỉnh:</div></td>
	  <td>
	  <select name="selTinh1" onChange="returnTinh();">
		<option value="0" selected>Chọn tỉnh</option>
		<%
		sql="Select * From Tinh"
		set rspr=Server.CreateObject("ADODB.Recordset")
		rspr.open sql,Con,1 
		Do while not rspr.eof
		%>
		<option value="<%=rspr("ProvinceID")%>" <%if ProvinceID = rspr("ProvinceID") then Response.Write("selected=""selected""") end if %>> <%=rspr("NameProvince")%></option>
		<%
		rspr.movenext
		Loop
		rspr.close
		%>
	  </select></td>
	  <td><span class="style8">(*)</span></td>
	</tr>
	<tr>
	  <td height="27"><div align="right">Huyện:</div></td>
	  <td class="CTxtContent">
		<select name="selHuyen1">
		  <option value="0" selected>Chọn Huyện</option>
		  <%

		sql="Select * From Huyen where ProvinceID='"&ProvinceID&"'"
		set rspr=Server.CreateObject("ADODB.Recordset")
		rspr.open sql,Con,1 
		iHuyen = 0
		Do while not rspr.eof
		%>
		  <option value="<%=rspr("DistrictID")%>" <%if DistrictID = rspr("DistrictID") then Response.Write("selected=""selected""") end if %>><%=rspr("NameDistrict")%></option>
		  <%
		iHuyen = iHuyen+ 1
		rspr.movenext
		Loop
		rspr.close
		%>
		</select>                              </td>
	  <td><span class="style8">(*)</span></td>
	</tr>
	<tr>
	  <td height="27"><div align="right"><span class="style7">Địa chỉ : </span></div></td>
	  <td><input name="f_diachi" id="f_diachi" size="35" class="CTextBoxUnder" value="<%=diachi%>">                              </td>
	  <td>&nbsp;</td>
	</tr>
	<tr>
	  <td height="27"><div align="right"><span class="style7">E-Mail : </span></div></td>
	  <td><input name="f_mail" type="text" id="f_mail" value="<%=Email%>" size="35" class="CTextBoxUnder"></td>
	  <td><span class="style8">(*)</span></td>
	</tr>
	<tr>
	  <td height="27"><div align="right"><span class="style7">Điện thoại cố định : </span></div></td>
	  <td><input name="f_tell" type="text" id="f_mobile2" size="35" class="CTextBoxUnder" value="<%=Tel%>">                              </td>
	  <td rowspan="2"><span class="style8">(*)</span></td>
	</tr>
	<tr>
	  <td height="29"><div align="right"><span class="style7">Điện thoại di động : </span></div></td>
	  <td><input name="f_mobile" type="text" id="f_mobile" size="35" class="CTextBoxUnder" value="<%=mobile%>">                              </td>
	</tr>
	<tr>
	  <td height="28"><div align="right"><span class="style7">Mật khẩu : </span></div></td>
	  <td><input name="f_pass" type="text" id="f_pass" value="<%=pass%>" size="35" class="CTextBoxUnder"></td>
	  <td><span class="style8">(*)</span></td>
	</tr>
	
	<tr>
	  <td colspan="3" align="center">
		  <input name="Submit" type="submit" id="Submit" onClick="CheckSubmit()" value="Sửa đổi">
			  <input type="reset" name="Submit2" value="Reset" >		</td>
	</tr>
</table>                    
</form>

</body>
</html>
<script language="javascript">
function returnTinh()
{
	document.form_reg.action = 'customer_edit.asp?reTinh=1';	
	document.form_reg.submit();
}
</script>
<script language="JavaScript">

function CheckSubmit()
{
	
	if (document.form_reg.f_cmnd.value == '')
	{
		alert('Xin hãy nhập chứng minh nhan dan');
		document.form_reg.f_cmnd.focus();
		return;
	}
	var cmnd=document.form_reg.f_cmnd.value
	if (cmnd < 100000 || cmnd > 999999999999)
	{
	alert('chứng minh nhân dân nhập không hợp lệ');
		document.form_reg.f_cmnd.focus();
		return;
	}

	var email=document.form_reg.f_mail.value
	var i=email.indexOf("@")
	var j=email.indexOf(".")
	if (i<2 || j<4)
	{
	alert('E-Mail nhập không hợp lệ');
		document.form_reg.f_mail.focus();
		return;
	}
	if (document.form_reg.f_mail.value=='')
	{
		alert('Dia chi Email khong hop le!');
		document.form_reg.f_mail.focus();
		return;
	}

	if (document.form_reg.f_pass.value==''||document.form_reg.f_pass2.value=='')
	{
		alert('Xin hãy nhập cả hai mật khẩu!');
		document.form_reg.f_pass.focus();
		return;
	}
	if (document.form_reg.f_pass.value!=document.form_reg.f_pass2.value)
	{
		alert('Mật khẩu gõ lại phải giống mật khẩu trên');
		document.form_reg.f_pass.focus();
		return;
	}
	
	if (document.form_reg.f_name.value == '')
	{
		alert('Xin hay nhap ho ten day du!');
		document.form_reg.f_name.focus();
		return;
	}

	if (document.form_reg.f_quequan.value == '')
	{
		alert('Xin hay nhap que quan!');
		document.form_reg.f_quequan.focus();
		return;
	}
	if (document.form_reg.f_tell.value=='' && document.form_reg.f_mobile.value=='')
	{
		alert('Xin hay nhap dien thoai!');
		document.form_reg.f_tell.focus();
		return;
	}
	
	document.form_reg.submit();
}
</script>