f<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/md5.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<%
f_permission = administrator(false,session("user"),"m_user")
txtGuide=""
if f_permission <= 2 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
if 	request.Form("action")="Insert" then
	booError=False
	UserName=Trim(replace(request.Form("UserName"),"'","''"))
	IDNhanVien=Trim(replace(request.Form("SelStaff"),"'","''"))
	UserPwd=Trim(replace(request.Form("UserPwd"),"'","''"))
	if CheckUserExist(Username)<>0 then
		sUserName="Đang sử dụng"
		booError=True
	end if
	If Userpwd<>"" and len(UserPwd)<6 then
		sUserPwd="> 6 ký tự"
		booError=True
	end if	
	
	if not booError then
		sql="insert into [User] (UserName,UserPwd,IDNhanVien) values "
		sql=sql & "(N'" & UserName & "'"
		sql=sql & ",'" & md5(UserPwd) & "'"
		sql=sql & ",'" & IDNhanVien & "')"

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.open sql,con,1
		
		m_editor		=	GetNumeric(Request.Form("m_editor"),0)
		m_order_output	=	GetNumeric(Request.Form("m_order_output"),0)
		m_order_input	=	GetNumeric(Request.Form("m_order_input"),0)
		m_user			=	GetNumeric(Request.Form("m_user"),0)
		m_customer		=	GetNumeric(Request.Form("m_customer"),0)
		m_report		=	GetNumeric(Request.Form("m_report"),0)
		m_accounting	=	GetNumeric(Request.Form("m_accounting"),0)
		m_sys			=	GetNumeric(Request.Form("m_sys"),0)
		m_human			=	GetNumeric(Request.Form("m_human"),0)
		m_work		 	=	GetNumeric(Request.Form("m_work"),0)
		m_cod			=	GetNumeric(Request.Form("m_cod"),0)
		m_sale			=	GetNumeric(Request.Form("m_sale"),0)
		m_ads			=	GetNumeric(Request.Form("m_ads"),0)
		m_faq			=	GetNumeric(Request.Form("m_faq"),0)
		adm				=	GetNumeric(Request.Form("adm"),0)
		
		sql="insert into UserDistribution (UserName, CategoryID, User_role, m_editor, m_order_output, m_order_input, m_user, m_customer, m_report, m_accounting, m_sys, m_human, m_work, m_cod, m_sale, m_ads, m_faq, adm) values "
		sql=sql & "(N'" & UserName & "'"
		sql=sql & "," & 0
		sql=sql & ",'ed'"
		sql=sql & ",'"&m_editor&"'"
		sql=sql & ",'"&m_order_output&"'"
		sql=sql & ",'"&m_order_input&"'"
		sql=sql & ",'"&m_user&"'"
		sql=sql & ",'"&m_customer&"'"
		sql=sql & ",'"&m_report&"'"
		sql=sql & ",'"&m_accounting&"'"
		sql=sql & ",'"&m_sys&"'"
		sql=sql & ",'"&m_human&"'"
		sql=sql & ",'"&m_work&"'"
		sql=sql & ",'"&m_cod&"'"
		sql=sql & ",'"&m_sale&"'"
		sql=sql & ",'"&m_ads&"'"
		sql=sql & ",'"&m_faq&"'"					
		sql=sql & ",'"&adm&"')"
		Response.Write("Hệ thống đã cập nhật thành công, mời nhấn OK để tiếp tục")
		rs.open sql,con,1
		set rs=nothing
	
		response.Write "<script language=""JavaScript"">" & vbNewline &_
			"<!--" & vbNewline &_
				"alert('Đã cập nhật');" & vbNewline &_
				"window.opener.location.reload();" & vbNewline &_
				"window.close();" & vbNewline &_
			"//-->" & vbNewline &_
			"</script>" & vbNewline
	end if
elseif request.Form("action")="update" then
	booError=False
	UserName=Trim(replace(request.Form("UserName"),"'","''"))
	IDNhanVien=Trim(replace(request.Form("SelStaff"),"'","''"))
	UserPwd=Trim(replace(request.Form("UserPwd"),"'","''"))
	If Userpwd<>"" and len(UserPwd)<6 then
		sUserPwd="> 6 ký tự"
		booError=True
	end if		
	if not booError then
		sql="update [User] set IDNhanVien='" & IDNhanVien & "'"
		if UserPwd <> "" then	
		sql=sql+ ",UserPwd='" & md5(UserPwd) & "'"		
		end if
		sql=sql+"where UserName='"&UserName&"'"
		
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.open sql,con,1
		
		m_editor		=	GetNumeric(Request.Form("m_editor"),0)
		m_order_output	=	GetNumeric(Request.Form("m_order_output"),0)
		m_order_input	=	GetNumeric(Request.Form("m_order_input"),0)
		m_out_store		=	GetNumeric(Request.Form("m_out_store"),0)
		m_user			=	GetNumeric(Request.Form("m_user"),0)
		m_customer		=	GetNumeric(Request.Form("m_customer"),0)
		m_report		=	GetNumeric(Request.Form("m_report"),0)
		m_accounting	=	GetNumeric(Request.Form("m_accounting"),0)
		m_sys			=	GetNumeric(Request.Form("m_sys"),0)
		m_human			=	GetNumeric(Request.Form("m_human"),0)
		m_work		 	=	GetNumeric(Request.Form("m_work"),0)
		m_cod			=	GetNumeric(Request.Form("m_cod"),0)
		m_sale			=	GetNumeric(Request.Form("m_sale"),0)
		m_ads			=	GetNumeric(Request.Form("m_ads"),0)
		m_faq			=	GetNumeric(Request.Form("m_faq"),0)
		adm				=	GetNumeric(Request.Form("adm"),0)
		
		sql="update UserDistribution set m_editor='"&m_editor&"', m_order_output='"&m_order_output&"', m_order_input='"&m_order_input&"', m_out_store='"& m_out_store &"', m_user='"&m_user&"', m_customer='"&m_customer&"', m_report='"& m_report &"', m_accounting='"&m_accounting&"', m_sys='"&m_sys&"', m_human='"&m_human&"', m_work='"&m_work&"', m_cod='"&m_cod&"', m_sale='"&m_sale&"', m_ads='"&m_ads&"', m_faq='"&m_faq&"', adm ='"&adm&"' where UserName='"&UserName&"'"
		Response.Write("Hệ thống đã cập nhật thành công, mời nhấn OK để tiếp tục")
		rs.open sql,con,1
		set rs=nothing
	
		response.Write "<script language=""JavaScript"">" & vbNewline &_
			"<!--" & vbNewline &_
				"alert('Đã cập nhật');" & vbNewline &_
				"window.opener.location.reload();" & vbNewline &_
				"window.close();" & vbNewline &_
			"//-->" & vbNewline &_
			"</script>" & vbNewline
	end if
end if

if Request.QueryString("act")="edit" then
	UserName=Request.QueryString("param")
	UserName=Replace(UserName,"'","''")
	
	sql="SELECT UserDistribution.UserName, IDNhanVien, m_editor, m_order_output, m_order_input, m_out_store,m_user, m_customer, m_report, m_accounting, m_sys, m_human, m_work, m_cod, m_sale, m_ads, m_faq, adm FROM [User] INNER JOIN UserDistribution ON [User].UserName =UserDistribution.UserName where [User].UserName='"&UserName&"'"			
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
	if not rs.eof then
		IDNhanVien		=	rs("IDNhanVien")
		
		m_editor		=	rs("m_editor")
		m_order_output	=	rs("m_order_output")
		m_order_input	=	rs("m_order_input")
		m_out_store		=	rs("m_out_store")
		m_user			=	rs("m_user")
		m_customer		=	rs("m_customer")
		m_report		=	rs("m_report")
		m_accounting	=	rs("m_accounting")
		m_sys			=	rs("m_sys")
		m_human			=	rs("m_human")
		m_work		 	=	rs("m_work")
		m_cod			=	rs("m_cod")
		m_sale			=	rs("m_sale")
		m_ads			=	rs("m_ads")
		m_faq			=	rs("m_faq")
		adm				=	rs("adm")	
	end if
	action	=	"update"
else
	action	=	"Insert"	
end if
%>
<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="../../css/styles.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style57 {color: #FF0000}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="fNew" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>?<%=Request.ServerVariables("QUERY_STRING")%>">
  <table width="100%" border="0" cellpadding="2" cellspacing="2" class="CTxtContent">
    <!--<tr align="center" valign="middle"> 
      <td height="40" colspan="2" valign="middle"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Tạo User mới</strong></font></td>
    </tr>-->
    <tr> 
      <td align="left"><font size="2" face="Arial, Helvetica, sans-serif">Tên 
        truy nhập:</font></td>
      <td align="left">
	  <%if Request.QueryString("act")="edit" then%>
	  <%=UserName%>
	   <input name="UserName" type="hidden" id="UserName" size="30" maxlength="30" value="<%=UserName%>">
	  <%else%>
	   <input name="UserName" type="text" id="UserName" size="30" maxlength="30" value="">
	   <%end if%>
        <font size="1" face="Arial, Helvetica, sans-serif">(<font color="#FF0000"><strong>*<%=sUserName%></strong></font>)</font></td>
    </tr>
    
    <tr> 
      <td align="left"><font size="2" face="Arial, Helvetica, sans-serif">Họ tên:</font></td>
      <td align="left">
	  <% call SelectNhanVien("SelStaff",IDNhanVien,6,0,0)%>	 	  </td>
    </tr>
    
    <tr> 
      <td align="left"><font size="2" face="Arial, Helvetica, sans-serif">Mật 
        khẩu:</font></td>
      <td align="left"><input name="UserPwd" type="text" id="UserPwd" size="30" maxlength="30">
        <font size="1" face="Arial, Helvetica, sans-serif">(<font color="#FF0000"><strong>*<%=sUserPwd%></strong></font>)</font></td>
    </tr>
    
    

    <tr>
      <td  colspan="2" height="35" valign="bottom" style="display:none;">
	  <font class="CTieuDeNho style57"> Thiết lập các quyền truy cập</font>
	  <br>
<table width="100%" border="1" cellpadding="1" cellspacing="1" class="CTxtContent">
<tr>
  <td valign="top" bgcolor="#FFFFFF" align="left" class="CTxtContent">
	
	<span class="CTieuDeNhoNho">Biên tập viên</span><br>
	<input name="m_editor" type="radio" id="m_editor" value="1" <%if m_editor = 1 then %> checked <%end if%> > 
	Tin tức &nbsp;&nbsp;&nbsp;
	<input name="m_editor" type="radio" id="m_editor" value="2" <%if m_editor = 2 then %> checked <%end if%> > 
	Sách &nbsp;&nbsp; 
    <input name="m_editor" type="radio" id="m_editor" value="3" <%if m_editor = 3 then %> checked <%end if%> > 
    Sản phẩm hitech 
&nbsp;
	<input name="m_editor" type="radio" id="m_editor" value="4" <%if m_editor = 4 then %> checked <%end if%> > 
	Xóa &nbsp;&nbsp;&nbsp;
	<input name="m_editor" type="radio" id="m_editor" value="0" <%if m_editor = 0 then %> checked <%end if%> >	 Không            </td> 
	<td valign="top" bgcolor="#FFFFFF" align="left" class="CTxtContent">
	<span class="CTieuDeNhoNho">Quản lý đơn hàng </span><br>
	<input name="m_order_output" type="radio" id="m_order_output" value="1" <%if m_order_output = 1 then %> checked <%end if%> >
	Xem đơn 
	&nbsp;&nbsp;&nbsp;
	<input name="m_order_output" type="radio" id="m_order_output" value="2" <%if m_order_output = 2 then %> checked <%end if%> > Sửa &nbsp;&nbsp;&nbsp;
	<input name="m_order_output" type="radio" id="m_order_output" value="3" <%if m_order_output = 3 then %> checked <%end if%> > Xóa &nbsp;&nbsp;&nbsp;
	<input name="m_order_output" type="radio" id="m_order_output" value="0" <%if m_order_output = 0 then %> checked <%end if%> >	 Không</td>
</tr>
<tr>
  <td valign="top" bgcolor="#FFFFFF" align="left" class="CTxtContent">
	
	<span class="CTieuDeNhoNho">Quản lý nhập hàng </span><br>
	<input name="m_order_input" type="radio" id="m_order_input" value="1" <%if m_order_input = 1 then %> checked <%end if%> >
	Xem-ký
	&nbsp;&nbsp;&nbsp;
	<input name="m_order_input" type="radio" id="m_order_input" value="2" <%if m_order_input = 2 then %> checked <%end if%> > 
	Thêm - sửa &nbsp;&nbsp;&nbsp;
	<input name="m_order_input" type="radio" id="m_order_input" value="3" <%if m_order_input = 3 then %> checked <%end if%> > Xóa &nbsp;&nbsp;&nbsp;
	<input name="m_order_input" type="radio" id="m_order_input" value="0" <%if m_order_input = 0 then %> checked <%end if%> >	 Không</td> 
	<td valign="top" bgcolor="#FFFFFF" align="left" class="CTxtContent"><span class="CTieuDeNhoNho">Quản lý xuất kho </span><br>
	  <input name="m_out_store" type="radio" id="radio" value="1" <%if m_out_store = 1 then %> checked <%end if%> >
	  Cho phép 
	  &nbsp;&nbsp;&nbsp;
      <input name="m_out_store" type="radio" id="radio4" value="0" <%if m_out_store = 0 then %> checked <%end if%> >
Không</td>
</tr>
<tr>
  <td valign="top" bgcolor="#FFFFFF" align="left" class="CTxtContent">
	
	<span class="CTieuDeNhoNho">Quản lý khách hàng </span><br>
	<input name="m_customer" type="radio" id="m_customer" value="1" <%if m_customer = 1 then %> checked <%end if%> > Cơ bản &nbsp;&nbsp;&nbsp;
	<input name="m_customer" type="radio" id="m_customer" value="2" <%if m_customer = 2 then %> checked <%end if%> > Chi tiết &nbsp;&nbsp;&nbsp;
	<input name="m_customer" type="radio" id="m_customer" value="3" <%if m_customer = 3 then %> checked <%end if%> > 
	Gửi email &nbsp;&nbsp;&nbsp;
	<input name="m_customer" type="radio" id="m_customer" value="0" <%if m_customer = 0 then %> checked <%end if%> >	 Không</td> 
	<td valign="top" bgcolor="#FFFFFF" align="left" class="CTxtContent">
	<span class="CTieuDeNhoNho">Quản lý thống kê </span><br>
	<input name="m_report" type="radio" id="m_report" value="1" <%if m_report = 1 then %> checked <%end if%> >
	Chỉ thống kê dữ liệu cá nhân &nbsp;&nbsp;&nbsp;
	<input name="m_report" type="radio" id="m_report" value="2" <%if m_report = 2 then %> checked <%end if%> >
	Chuyên sâu	&nbsp;&nbsp;&nbsp;
	<input name="m_report" type="radio" id="m_report" value="0" <%if m_report = 0 then %> checked <%end if%> >	 Không </td>
</tr>
<tr>
  <td valign="top" bgcolor="#FFFFFF" align="left" class="CTxtContent">
	
	<span class="CTieuDeNhoNho">Quản lý kế toán chứng từ </span><br>
	<input name="m_accounting" type="radio" id="m_accounting" value="1" <%if m_accounting = 1 then %> checked <%end if%> > 
	Nhập liệu &nbsp;&nbsp;&nbsp;
	<input name="m_accounting" type="radio" id="m_accounting" value="2" <%if m_accounting = 2 then %> checked <%end if%> >
	Tổng hợp &nbsp;&nbsp;
	<input name="m_accounting" type="radio" id="m_accounting" value="0" <%if m_accounting = 0 then %> checked <%end if%> >	 Không</td> 
	<td valign="top" bgcolor="#FFFFFF" align="left" class="CTxtContent">
	<span class="CTieuDeNhoNho">Ưu tiên   hệ thống </span><br>
	<input name="m_sys" type="radio" id="m_sys" value="1" <%if m_sys = 1 then %> checked <%end if%> >
	Có
	&nbsp;&nbsp;&nbsp;
	<input name="m_sys" type="radio" id="m_sys" value="0" <%if m_sys = 0 then %> checked <%end if%> >	 Không</td>
</tr>
<tr>
  <td valign="top" bgcolor="#FFFFFF" align="left" class="CTxtContent">
	
	<span class="CTieuDeNhoNho">Quản trị nhân sự </span><br>
    <input name="m_human" type="radio" id="radio" value="1" <%if m_human = 1 then %> checked <%end if%> >	
    Chỉ xem User&nbsp;&nbsp;&nbsp;
	<input name="m_human" type="radio" id="m_human" value="2" <%if m_human = 2 then %> checked <%end if%> >
	Quản lý&nbsp;&nbsp;&nbsp;
	<input name="m_human" type="radio" id="m_human" value="0" <%if m_human = 0 then %> checked <%end if%> >	 Không</td> 
	<td valign="top" bgcolor="#FFFFFF" align="left" class="CTxtContent">
	<span class="CTieuDeNhoNho">Quản trị công việc </span><br>
	<input name="m_work" type="radio" id="m_work" value="1" <%if m_work = 1 then %> checked <%end if%> > Thêm &nbsp;&nbsp;&nbsp;
	<input name="m_work" type="radio" id="m_work" value="2" <%if m_work = 2 then %> checked <%end if%> > Sửa &nbsp;&nbsp;&nbsp;
	<input name="m_work" type="radio" id="m_work" value="3" <%if m_work = 3 then %> checked <%end if%> > Xóa &nbsp;&nbsp;&nbsp;
	<input name="m_work" type="radio" id="m_work" value="0" <%if m_work = 0 then %> checked <%end if%> >	 Không</td>
</tr>
<tr>
  <td valign="top" bgcolor="#FFFFFF" align="left" class="CTxtContent">
	
	<span class="CTieuDeNhoNho">Quản lý COD </span><br>
	<input name="m_cod" type="radio" id="m_cod" value="1" <%if m_cod = 1 then %> checked <%end if%> > 
	Tiếp nhận phiếu &nbsp;&nbsp;&nbsp;
	<input name="m_cod" type="radio" id="m_cod" value="2" <%if m_cod = 2 then %> checked <%end if%> >
	Kiểm soát
	&nbsp;&nbsp;&nbsp;&nbsp;
	<input name="m_cod" type="radio" id="m_cod" value="0" <%if m_cod = 0 then %> checked <%end if%> >	 Không</td> 
	<td valign="top" bgcolor="#FFFFFF" align="left" class="CTxtContent">
	<span class="CTieuDeNhoNho">Quản trị khuyến mãi </span><br>
	<input name="m_sale" type="radio" id="m_sale" value="1" <%if m_sale = 1 then %> checked <%end if%> > Thêm &nbsp;&nbsp;&nbsp;
	<input name="m_sale" type="radio" id="m_sale" value="2" <%if m_sale = 2 then %> checked <%end if%> > Sửa &nbsp;&nbsp;&nbsp;
	<input name="m_sale" type="radio" id="m_sale" value="3" <%if m_sale = 3 then %> checked <%end if%> > Xóa &nbsp;&nbsp;&nbsp;
	<input name="m_sale" type="radio" id="m_sale" value="0" <%if m_sale = 0 then %> checked <%end if%> >	 Không</td>
</tr>
<tr>
  <td valign="top" bgcolor="#FFFFFF" align="left" class="CTxtContent">
	
	<span class="CTieuDeNhoNho">Quản lý quảng cáo </span><br>
	<input name="m_ads" type="radio" id="m_ads" value="1" <%if m_ads = 1 then %> checked <%end if%> > Thêm &nbsp;&nbsp;&nbsp;
	<input name="m_ads" type="radio" id="m_ads" value="2" <%if m_ads = 2 then %> checked <%end if%> > Sửa &nbsp;&nbsp;&nbsp;
	<input name="m_ads" type="radio" id="m_ads" value="3" <%if m_ads = 3 then %> checked <%end if%> > Xóa &nbsp;&nbsp;&nbsp;
	<input name="m_ads" type="radio" id="m_ads" value="0" <%if m_ads = 0 then %> checked <%end if%> >	 Không</td> 
	<td valign="top" bgcolor="#FFFFFF" align="left" class="CTxtContent">
	<span class="CTieuDeNhoNho">Quản trị FAQ </span><br>
	<input name="m_faq" type="radio" id="m_faq" value="1" <%if m_faq = 1 then %> checked <%end if%> > Thêm &nbsp;&nbsp;&nbsp;
	<input name="m_faq" type="radio" id="m_faq" value="2" <%if m_faq = 2 then %> checked <%end if%> > Sửa &nbsp;&nbsp;&nbsp;
	<input name="m_faq" type="radio" id="m_faq" value="3" <%if m_faq = 3 then %> checked <%end if%> > Xóa &nbsp;&nbsp;&nbsp;
	<input name="m_faq" type="radio" id="m_faq" value="0" <%if m_faq = 0 then %> checked <%end if%> >	 Không</td>
</tr>
<tr>
  <td valign="top" bgcolor="#FFFFFF" align="left" class="CTxtContent"><span class="CTieuDeNhoNho">Quản trị User </span><br>
    <input name="m_user" type="radio" id="m_user" value="1" <%if m_user = 1 then %> checked <%end if%> >
Thêm &nbsp;&nbsp;&nbsp;
<input name="m_user" type="radio" id="radio2" value="2" <%if m_user = 2 then %> checked <%end if%> >
Sửa &nbsp;&nbsp;&nbsp;
<input name="m_user" type="radio" id="m_user" value="3" <%if m_user = 3 then %> checked <%end if%> >
Xóa &nbsp;&nbsp;&nbsp;
<input name="m_user" type="radio" id="m_user" value="0" <%if m_user = 0 then %> checked <%end if%> >
Không</td>
  <td valign="top" bgcolor="#FFFFFF" align="left" class="CTxtContent"><span class="CTieuDeNhoNho">Kiểm soát mọi chức năng </span><br>
    <input name="adm" type="radio" id="adm" value="1" <%if adm = 1 then %> checked <%end if%> >
Có &nbsp;&nbsp;&nbsp;
<input name="adm" type="radio" id="adm" value="0" <%if adm = 0 then %> checked <%end if%> >
Không</td>
</tr>
	  </table>      </td>
    </tr>
    <tr> 
      <td align="center" colspan="2" height="35" valign="bottom"> <input type="submit" name="Submit" value="Cập nhật"> 
        <input type="button" name="Submit2" value="Đóng cửa sổ" onClick="javascript: window.close();"> 
		<input type="hidden" name="action" value="<%=action%>">      </td>
    </tr>
  </table>
</form>
</body>
</html>
