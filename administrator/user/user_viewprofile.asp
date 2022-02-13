<%session.CodePage=65001%>
<%	if Trim(session("user"))="" then
		response.Redirect("/administrator/default.asp")
		response.End()
	end if
%>	
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
	Dim rs
	set rs=Server.CreateObject("ADODB.Recordset")
%>
<HTML>
	<HEAD>
		<TITLE><%=PAGE_TITLE%></TITLE>
		<META http-equiv=Content-Type content="text/html; charset=utf-8">
		<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
        <script  type="text/javascript" src="/administrator/inc/common.js"></script>
        <link href="/administrator/inc/admstyle.css" type="text/css" rel="stylesheet">
        <link href="/administrator/skin/css/sweetalert.css" rel="stylesheet" />
        <link href="/administrator/skin/css/bootstrap.css" rel="stylesheet" />
        <script src="/administrator/skin/script/jquery-2.2.2.min.js"></script>
        <script src="/administrator/skin/script/ajax-asp.js"></script>
        <link href="/administrator/css/skin1.css" rel="stylesheet" />
	</HEAD>
<BODY leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
	Call header()
%>
<div class="container-fluid">
    <div class="col-md-2" style="background:#001e33;"><%call MenuVertical(0) %></div>
    <div class="col-md-10">
<table width="770" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr align="right" valign="top"> 
    <td height="25"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
      <a href="#changepass"> Đổi mật khẩu</a> | <a href="#editprofile">Sửa thông 
      tin cá nhân</a></strong></font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
  </tr>
</table>
<table width="96%" border="0" align="center" cellpadding="2" cellspacing="2">
  <tr align="center"> 
    <td width="40%" align="left"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Thông 
      tin cá nhân</strong></font></td>
    <td width="60%"><strong><font size="2" face="Arial, Helvetica, sans-serif">Các 
      quyền được cấp</font></strong></td>
  </tr>
  <%
  	sql="SELECT * from [USER] where username=N'" & session("user") & "'"
	rs.open sql,con,1
		CreationDate=ConvertTime(rs("CreationDate"))
		LastLoginDate=ConvertTime(rs("LastLoginDate"))
	rs.close
	set rs=nothing
  %>
  <tr> 
    <td><table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td width="5%">&nbsp;</td>
          <td><font size="2" face="Arial, Helvetica, sans-serif">Tên truy nhập:</font></td>
          <td><font size="2" face="Arial, Helvetica, sans-serif"><%=Session("user")%></font></td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
          <td><font size="2" face="Arial, Helvetica, sans-serif">Email:</font></td>
          <td><font size="2" face="Arial, Helvetica, sans-serif"><%=Email%></font></td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
          <td><font size="2" face="Arial, Helvetica, sans-serif">Tên đầy đủ:</font></td>
          <td><font size="2" face="Arial, Helvetica, sans-serif"><%=Session("fullname")%></font></td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
          <td><font size="2" face="Arial, Helvetica, sans-serif">Chức danh:</font></td>
          <td><font size="2" face="Arial, Helvetica, sans-serif"><%=Session("TitleUser")%></font></td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
          <td><font size="2" face="Arial, Helvetica, sans-serif">Phòng ban:</font></td>
          <td><font size="2" face="Arial, Helvetica, sans-serif"><%=Session("room")%></font></td>
        </tr>		
        <tr> 
          <td>&nbsp;</td>
          <td><font size="2" face="Arial, Helvetica, sans-serif">Khởi tạo:</font></td>
          <td><font size="2" face="Arial, Helvetica, sans-serif"><%=Day(CreationDate)%>/<%=Month(CreationDate)%>/<%=Year(CreationDate)%></font></td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
          <td><font size="2" face="Arial, Helvetica, sans-serif">Đăng nhập<br>
            lần cuối:</font></td>
          <td><font size="2" face="Arial, Helvetica, sans-serif">
			  <%=Hour(LastLoginDate)%><SUP>h</SUP><%=Minute(LastLoginDate)%>&nbsp;<%=Day(LastLoginDate)%>/<%=Month(LastLoginDate)%>/<%=Year(LastLoginDate)%>
		  </font></td>
        </tr>
      </table></td>
    <td rowspan="5" valign="top"> 
			<font class="CTxtContent">
					<font color="#FF0000">
					<%if GetNumeric(session("iBienTap"),0) = 1then%>
						- Biên tập viên<br>
					<%end if%>

					<%if GetNumeric(Session("iQLyDonHang"),0) = 1then%>
						- Quản lý đơn hàng<br>
					<%end if%>
					
					<%if GetNumeric(Session("iQLyNhapHang"),0) = 1then%>
						- Quản lý nhập hàng<br>
					<%end if%>
							
					<%if GetNumeric(Session("iQLyNhanVien"),0) = 1then%>
						- Quản lý nhân viên<br>
					<%end if%>
					
					<%if GetNumeric(Session("iQLyKhachHang"),0) = 1 then%>
						- Quản lý khách hàng<br>
					<%end if%>						
					
					<%if GetNumeric(Session("iQLyThongKe"),0) = 1then%>
						- Quản lý thống kê báo cáo<br>
					<%end if%>													

					<%if GetNumeric(Session("iQLyHeThong"),0) = 1then%>
						- Quản lý hệ thống<br>
					<%end if%>
					
					<%if GetNumeric(Session("iQLyKeToan"),0) = 1then%>
						- Quản lý kế toán<br>
					<%end if%>

					<%if GetNumeric(Session("iQuanTri"),0) = 1 then%>
						- Quản trị xóa dữ liệu<br>
					<%end if%>					
			</font>	
	</td>
  </tr>
</table>
         <table style="margin-block-start: 5rem; margin-left: 30px;">
        <tr> 
    <td class = "w3-margin-top" bgcolor="#808080"><a name="changepass"></a><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">&nbsp;<img src="/administrator/images/icon1.gif" width="7" height="7" align="absmiddle">&nbsp;Đổi 
      mật khẩu</font></strong></td>
  </tr>
    <tr> 
    <td><form style="padding: 2rem;" name="fChangePwd" method="post" action="user_changepassword.asp" onSubmit="myOpenWindow('user_changepassword.asp',300,150)" target="myWindowName">
        <table border="0" cellspacing="2" cellpadding="2" width="100%">
          <tr> 
            <td width="5%">&nbsp;</td>
            <td width="35%" align="left"><font size="2" face="Arial, Helvetica, sans-serif">Mật 
              khẩu cũ:&nbsp;</font></td>
            <td width="60%"> 
              <input name="oldpwd" type="password" id="oldpwd" size="23" maxlength="50">            </td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td align="left"><font size="2" face="Arial, Helvetica, sans-serif">Mật 
              khẩu mới:&nbsp;</font></td>
            <td><input name="newpwd" type="password" id="newpwd" size="23" maxlength="50"></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td align="left"><font size="2" face="Arial, Helvetica, sans-serif">Gõ 
              lại&nbsp;:</font></td>
            <td><input name="newpwdcon" type="password" id="newpwdcon" size="23" maxlength="50"></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td><input type="submit" name="submit" value="Thay đổi">
			<input type="hidden" name="action" value="password"></td>
          </tr>
        </table>
      </form></td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
    </table>
</div>
   
</div>
    
<%Call Footer()%>
</BODY>
</HTML>