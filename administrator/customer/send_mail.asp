<%session.CodePage=65001%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/include/func_tiny.asp" -->
<%
f_permission = administrator(false,session("user"),"m_order_output")
if f_permission < 2 then
	response.Redirect("/administrator/info.asp")
end if
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>XBOOK - SEND EMAIL</title>
<link href="../../css/styles.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<%
	img	="../../images/icons/icon_email1.gif"
	Title_This_Page="Khàng hàng -> Gửi email"
	Call header()
	Call Menu()

	
%>
<%
	iSendTT	=	GetNumeric(Request.QueryString("iSendTT"),0)
	SanPhamUser_ID=Request.QueryString("SanPhamUser_ID")
	sql="SELECT	TOP 1 * " &_
		"FROM  SanPhamUser " &_
		"WHERE     SanPhamUser_ID = '"& SanPhamUser_ID&"'"
	Set rs=server.CreateObject("ADODB.Recordset")
	rs.open sql,con,1
	if not rs.eof then
		SanPhamUser_Name=rs("SanPhamUser_Name")
		SanPhamUser_Email=rs("SanPhamUser_Email")
		SanPhamUser_Address=rs("SanPhamUser_Address")
		email	= SanPhamUser_Email
		giaoHang_name 	= 	rs("giaoHang_name")
		giaoHang_Email	= 	rs("giaoHang_Email")
		giaoHang_Address=	rs("giaoHang_Address")	
		SanPhamUser_Thoigian=rs("SanPhamUser_Thoigian")	
	end if
			
	strTemp = "ea4bb1a25bfd1b838a8a940d02c8d8ec"&SanPhamUser_ID&"d51dccca677d8049c356f8e5c830d7fc"
	linkinfo	=	"http://xbook.com.vn/ReportDH.asp?xboooooooo000000k="&strTemp
	munb=1000+SanPhamUser_ID
	strTemp	="XB"+CStr(munb)	
		
	if iSendTT = 1 then	
		subject	=	"Thong tin thanh toan tai xbook"
		if UCASE(giaoHang_Address) <> UCASE(SanPhamUser_Address) and trim(giaoHang_Address)<>"" then
			email	=	giaoHang_Email
			xungho	=	fXungHo(giaoHang_name)
			content	=	"<table align=""center"" width=""100%""  background=""http://xbook.com.vn/images/BGMain.gif""><tr><td><table  border=""0"" align=""center"" cellpadding=5 cellspacing=5  style=""border:#CCCCCC solid 1"" bgcolor=""#FFFFFF"" width=""500"" ><tr><td>XBOOK kính chào "&xungho&" "&giaoHang_name&"<br>XBOOK đã nhận được tiền thanh toán của "&xungho&" cho đơn hàng <a href="&linkinfo&" target=""_blank"">"&strTemp&"<a>. XBOOK sẽ sớm chuyển hàng tới "&fXungHo(SanPhamUser_Name)&" "&SanPhamUser_Name&"<br><br>Cám ơn "&xungho&" đã quan tâm và sử dụng dịch vụ của XBOOK <br>Chúc "&xungho&" mạnh khỏe hạnh phúc và thành công.<br>"
		else
			xungho	=	fXungHo(SanPhamUser_Name)			
			content	=	"<table align=""center"" width=""100%""  background=""http://xbook.com.vn/images/BGMain.gif""><tr><td><table  border=""0"" align=""center"" cellpadding=5 cellspacing=5  style=""border:#CCCCCC solid 1"" bgcolor=""#FFFFFF"" width=""500"" ><tr><td>XBOOK kính chào "&xungho&" "&SanPhamUser_Name&"<br>XBOOK đã nhận được tiền thanh toán của "&xungho&" cho đơn hàng <a href="&linkinfo&" target=""_blank"">"&strTemp&"<a> XBOOK sẽ sớm chuyển hàng tới "&xungho&"<br><br>Cám ơn "&xungho&" đã quan tâm và sử dụng dịch vụ của XBOOK <br>Chúc "&xungho&" mạnh khỏe hạnh phúc và thành công.<br>"
		end if
	elseif iSendTT = 2 then
		subject	=	"Thong tin thanh toan tai xbook"
		if UCASE(giaoHang_Address) <> UCASE(SanPhamUser_Address) and trim(giaoHang_Address)<>"" then
			email				=	giaoHang_Email
			SanPhamUser_Name	=	giaoHang_name
		end if
		xungho	=	fXungHo(SanPhamUser_Name)			
		content	=	"<table align=""center"" width=""100%""  background=""http://xbook.com.vn/images/BGMain.gif""><tr><td><table  border=""0"" align=""center"" cellpadding=5 cellspacing=5  style=""border:#CCCCCC solid 1"" bgcolor=""#FFFFFF"" width=""500"" ><tr><td>XBOOK kính chào "&xungho&" "&SanPhamUser_Name&"<br>XBOOK đã xác nhận thông tin của "&xungho&" cho đơn hàng <a href="&linkinfo&" target=""_blank"">"&strTemp&"<a><br>"&SanPhamUser_Thoigian&"<br> XBOOK mong sớm được chuyển hàng tới "&xungho&"<br><br>Cám ơn "&xungho&" đã quan tâm và sử dụng dịch vụ của XBOOK <br>Chúc "&xungho&" mạnh khỏe hạnh phúc và thành công.<br>"		
	 elseif iSendTT = 3 then
	 	subject	=	"XBOOK lien he thong tin"
	end if
%>
	<form id="form1" name="form1" method="post" action="send.asp">
      <table width="630" border="0" align="center">
        <tr>
          <td colspan="2" align="center" class="author"><div align="center"><br>
          SEND EMAIL<br>
          </div></td>
          </tr>
        <tr>
          <td colspan="2">&nbsp;</td>
        </tr>
        
        
        <tr>
          <td align="right" class="CTxtContent">Từ:</td>
          <td><input name="txt_form" type="text" id="txt_form" value="info@xbook.com.vn" size="30"></td>
        </tr>
        <tr>
          <td width="110" align="right" class="CTxtContent">Tới:</td>
          <td width="510"><textarea name="txt_mail" cols="60" rows="2"><%=email%></textarea></td>
        </tr>
        <tr>
          <td align="right" class="CTxtContent">Tiêu đề: </td>
          <td><input name="txt_subject" type="text" id="txt_subject" value="<%=subject%>" size="60"></td>
          </tr>
        
        <tr>
		 <td align="right" class="CTxtContent" valign="top">
		 Nội dung:          </td>
          <td class="CTxtContent" valign="top">
		  <%
		  	Signature	=	"<br><br> Lien he: <br>Sieu thi sach truc tuyen XBOOK <br>Website:<a href=""http://www.xbook.com.vn"" target=""_blank""> http://www.xbook.com.vn</a> <br> Tel: "& soDT &" <br>Email:info@xbook.com.vn<br>D/c: So 46/897 Pricei Phong - Hoang Mai - Ha Noi.</td></tr></table></td></tr></table><br>"			
			content	=	content+Signature
		  %>
            <textarea name="txt_content" cols="80" rows="20" id="txt_content" class="CTxtContent"><%=content%></textarea>    
			      </td>
          </tr>
        <tr>
          <td colspan="2"><table width="341" border="0" align="center">
            <tr>
              <td width="161"><div align="center">
                <input name="OK" type="submit" id="OK" value="   Gửi   " />
              </div></td>
              <td width="170"><div align="center">
                <input type="reset" name="Reset" value="Reset" />
              </div></td>
            </tr>
          </table></td>
          </tr>
      </table>
      <div align="center"></div>
	</form>
<script src='../inc/news.js'></script>
<script>VISUAL=4; FULLCTRL=1;</script>
<script src='../js/quickbuild.js'></script>
<script>changetoIframeEditor(document.forms[0].txt_content)</script>
</body>
</html>

