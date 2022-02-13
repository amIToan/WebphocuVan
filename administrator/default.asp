<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<%
	session.Abandon()
	sError=Request.QueryString("sError")
%>
<HTML>
<HEAD>
	<TITLE><%=PAGE_TITLE%></TITLE>
	<META http-equiv=Content-Type content="text/html; charset=utf-8">
    <link href="../css/style_show_image.css" rel="stylesheet" />
    <link href="../css/styles.css" rel="stylesheet" />
    <link href="../css/CommonSite.css" rel="stylesheet" />
</HEAD>
<BODY leftMargin=0 topMargin=0 onLoad="javascript: document.fLogin.uid.focus();">
<br>
<br>
<br>
    <div style="width:1000px;margin:auto;">
        <table class="CTxtContent"  style="margin:auto;">
            <tr>
                <td style="vertical-align:middle;">
                    <div class="box-logo-top" style="opacity:0.7;"> <img src="/images/logo/<%=Logo%>"  style="padding:5px;width:128px" /> </div>
                    <img src="images/xwork.png" class="box-border-no-fix" />
                </td>
                <td style="padding:30px;">
                    <form action="login.asp" method="post" name="fLogin">
                         <img src="/Images/Icons/check.png" />
                        <div class="CTieuDeNho"><%=company%></div><br />
                        <hr />
                        <div style="margin:auto;">
                            Tên Đăng Nhập:<br />
                            <input name="uid" type="text" id="uid" size="35" style="height:30px"><br /><br />
                            Mật Khẩu:<br />
                            <input name="pwd" type="password" id="pwd" size="35" style="height:30px">
                        </div>
                        <div><br />
                            <%if sError="Invalid" then%>
                            <img src="images/warning.jpg" width="34" height="34" align="middle">
                              <div  style="color:#f00;">User hoặc password không đúng</div>
                            <%elseif sError="Inactive" then%>
                            <div style="color:#f00;"> Đăng nhập không thành công do hết hạn hợp đồng</div>
                            <%else%>
                            <div  class="CTieuDeNhoNho"> Lưu ý: Hạn chế được lưu password tại máy tính vãng lai.<br /></div>
                            <%End if%>
                        </div>
                        <div style="padding:30px;">
                            <input type="submit" name="Submit" value="  Bắt đầu  " style="background-color:#055F91; width:150px;height:30px; color:#fff;">
                        </div>                   
                    </form>
                </td>
            </tr>
        </table>
         <%Call Footer()%>
    </div>
   


 </BODY>
</HTML>