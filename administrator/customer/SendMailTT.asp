<%session.CodePage=65001%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<%
f_permission = administrator(false,session("user"),"m_customer")
if f_permission < 3 then
	response.Redirect("/administrator/info.asp")
end if
%>
<script src="../../ckeditor/ckeditor.js"></script>
<script src="../../ckfinder/ckfinder.js"></script>

<html>
	<head>
	<title><%=PAGE_TITLE%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
	<link href="../../css/styles.css" rel="stylesheet" type="text/css">
	</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
	img	="../../images/icons/icon_email1.gif"
	Title_This_Page="Khách hàng -> Soạn và gửi email khách hàng"
	Call header()

%>

<table width="950" border="0" align="center" cellpadding="0" cellspacing="0" >
  <tr>
    <td Class="CtieuDe" align="center">SOẠN MAIL</td>
  </tr>
  <tr>
    <td ><br>
	<form id="form1" name="form1" method="post" action="sendTT.asp">
	  <table width="95%" border="0" align="center" cellpadding="1" cellspacing="1" class="CTxtContent" style="text-align:left;">
        <tr>
          <td width="27%" valign="top" style="<%=setStyleBorder(0,1,0,0)%>">
		  <span class="CTieuDeNho">Danh sách:</span><br>
		<%
		set arEmail		=	Session("arEmail")
		set arName		=	Session("arName")
		ST = 0
        if Ubound(arEmail.items) < 500 then
		    For Each Key in arEmail	
			    Ten			=	arName(Key)
                if Ten <> "" then
			        Response.Write(ST+1&".<b>"&ten&"</b><br>"	)
                else
                    Response.Write(ST+1&"."& arEmail(Key) &"<br>")
                end if
			    ST = ST + 1
		    next
        else
            Response.Write("Email list is very long so can not listed<br>We will send "& UBound(arEmail.items) &" Emails")  
            ST=UBound(arEmail.items)
        end if
		%>
		<hr />
		      Total: <%=ST%>		  </td>
          <td width="73%"  valign="top">
            Symbol name:
            <input name="txtHoTen" type="text" class="CTextBoxUnder" id="txtHoTen" size="15" maxlength="20"> 
              <span class="CSubTitle">Họ và Tên</span><br>
              Email gửi: 
                <input name="txtForm" type="text" class="CTextBoxUnder" id="txtForm" size="25"> Mật khẩu: 
                <input name="txtPassword" type="password" class="CTextBoxUnder" id="txtPassword" size="20">
              <br>
              Tên email: 
                <input name="txtname" type="text" class="CTextBoxUnder" id="txtname" size="45">
              <hr />
              Server Email: 
                <input name="txtserverIP" type="text" class="CTextBoxUnder" id="txtserverIP" size="45"><br />
              
               <br> 
              <hr />
              Tiêu đề: 
              <input name="txtTieuDe" type="text" class="CTextBoxUnder" id="txtTieuDe" size="80">
              <br>
            <br>
            Nội dung:<br><br>
            <textarea name="txtNoiDung" id="txtNoiDung" style="height:600px; width:800px" ></textarea></td>
        </tr>
        
        <tr>
          <td align="center">&nbsp;</td>
          <td align="center">&nbsp;</td>
        </tr>
        <tr>
          <td colspan="2" align="center" style="<%=setStyleBorder(0,0,1,0)%>"><br>
				  <input type="submit" name="Submit" value=" Gửi đi "><br></td>
          </tr>
      </table>	
	  </form>
	  </td>
  </tr>
  <tr>
    <td ></td>
  </tr>
</table>
</body>
</html>

</script>
    <script type="text/javascript">
        ///  CKEDITOR.replace('bodyx'); 
        CKEDITOR.replace('txtNoiDung');
        var editor = CKEDITOR.replace('txtNoiDung');
        CKFinder.setupCKEditor(editor, '/ckfinder/');
</script>