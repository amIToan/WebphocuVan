<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Call PhanQuyen("QLyNhanVien")%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
	username=Request.QueryString("param")
	username=trim(replace(username,"'","''"))
	strcat=""
	if request.Form("action")="Insert"  then
		set rs=server.CreateObject("ADODB.Recordset")
		iBienTap		= GetNumeric(Request.Form("iBienTap"),0)
		iQLyDonHang 	= GetNumeric(Request.Form("iQLyDonHang"),0)
		iQLyNhapHang	= GetNumeric(Request.Form("iQLyNhapHang"),0)
		iQLyNhanVien	= GetNumeric(Request.Form("iQLyNhanVien"),0)
		iQLyKhachHang	= GetNumeric(Request.Form("iQLyKhachHang"),0)
		iQLyThongKe		= GetNumeric(Request.Form("iQLyThongKe"),0)
		iQLyKeToan		= GetNumeric(Request.Form("iQLyKeToan"),0)
		iQLyHeThong		= GetNumeric(Request.Form("iQLyHeThong"),0)
		iQuanTri		= GetNumeric(Request.Form("iQuanTri"),0)

		
		sql = 	"update UserDistribution set "
		sql	=	sql+"  BienTap 		= "&iBienTap
		sql	=	sql+" ,QLyDonHang 	= "&iQLyDonHang		
		sql	=	sql+" ,QLyNhapHang 	= "&iQLyNhapHang		
		sql	=	sql+" ,QLyNhanVien	= "&iQLyNhanVien
		sql	=	sql+" ,QLyKhachHang	= "&iQLyKhachHang		
		sql	=	sql+" ,QLyThongKe	= "&iQLyThongKe		
		sql	=	sql+" ,QLyKeToan	= "&iQLyKeToan		
		sql	=	sql+" ,QLyHeThong 	= "&iQLyHeThong
		sql	=	sql+" ,Adm 	= "&iQuanTri
		sql	=	sql+" where UserName ='"& username &"'"
		Response.Write(sql)
		rs.open sql,con,1

		set rs=nothing
		
		response.Write "<script language=""JavaScript"">" & vbNewline &_
		"<!--" & vbNewline &_
		"window.opener.location.reload();" & vbNewline &_
		"window.close();" & vbNewline &_
		"//-->" & vbNewline &_
		"</script>" & vbNewline
	end if

%>
<html>
<head>
<title><%=PAGE_TITLE%></title></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="../../css/styles.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr align="center" valign="middle"> 
    <td height="60"><font size="3" face="Arial, Helvetica, sans-serif">Cấp quyền 
      cho User</font><font size="3" face="Arial, Helvetica, sans-serif"><strong><br>
      <font size="2">(<%=Username%>) </font></strong></font></td>
  </tr>
  <tr> 
    <td>
	<form name="fRoles" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>?<%=Request.ServerVariables("QUERY_STRING")%>">
	<%
		sql="Select * FROM UserDistribution where UserName='"&username&"'"
		set rs=server.CreateObject("ADODB.Recordset")
		rs.open sql,con,1
		if not rs.eof then
			iBienTap		= GetNumeric(rs("BienTap"),0)
			iQLyDonHang 	= GetNumeric(rs("QLyDonHang"),0)
			iQLyNhapHang	= GetNumeric(rs("QLyNhapHang"),0)
			iQLyNhanVien	= GetNumeric(rs("QLyNhanVien"),0)
			iQLyKhachHang	= GetNumeric(rs("QLyKhachHang"),0)
			iQLyThongKe		= GetNumeric(rs("QLyThongKe"),0)
			iQLyKeToan		= GetNumeric(rs("QLyKeToan"),0)
			iQLyHeThong		= GetNumeric(rs("QLyHeThong"),0)
			iQuanTri		= GetNumeric(rs("Adm"),0)
		end if
		set rs=nothing
	%>
	<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#006699">
        <tr> 
          <td width="20%" bgcolor="#FFFFFF"><div align="center"></div></td>
        </tr>
        <tr> 
            <td valign="top" bgcolor="#FFFFFF" align="left" class="CTxtContent">
              <p>
                <input name="iBienTap" type="checkbox" id="iBienTap" value="1" <%if iBienTap <> 0 then %> checked <%end if%> >
              Biên tập viên</p>
              <p>
                <input name="iQLyDonHang" type="checkbox" id="iQLyDonHang" value="1" <%if iQLyDonHang <> 0 then %> checked <%end if%>>
             Quản lý đơn hàng</p>
              <p>
                <input name="iQLyNhapHang" type="checkbox" id="iQLyNhapHang" value="1" <%if iQLyNhapHang <>0 then %> checked <%end if%>>
             Quản lý nhập hàng</p>
              <p> 
                <input name="iQLyNhanVien" type="checkbox" id="iQLyNhanVien" value="1" <%if iQLyNhanVien <> 0 then %> checked <%end if%>>
                Quản lý nhân viên</p>
              <p>
                <input name="iQLyKhachHang" type="checkbox" id="iQLyKhachHang" value="1" <%if iQLyKhachHang <> 0 then %> checked <%end if%>>
                Quản lý khách hàng</p>
              <p>
                <input name="iQLyThongKe" type="checkbox" id="iQLyThongKe" value="1" <%if iQLyThongKe <> 0 then %> checked <%end if%>>
                Quản lý thống kê </p>
              <p>
                <input name="iQLyKeToan" type="checkbox" id="iQLyKeToan" value="1" <%if iQLyKeToan <> 0 then %> checked <%end if%>>
                Quản lý kế toán               </p>
              <p>
                <input name="iQLyHeThong" type="checkbox" id="iQLyHeThong" value="1" <%if iQLyHeThong <> 0 then %> checked <%end if%>>
Quản lý hệ thống </p>
              <p>
                <input name="iQuanTri" type="checkbox" id="iQuanTri" value="1" <%if iQuanTri <> 0 then %> checked <%end if%>>
              Quản trị</p>
              <p>
			   <input type="hidden" name="action" value="Insert">
               <input type="submit" name="Submit" value=" Cấp quyền ">
                <br>
                (<font color="#FF0000"><strong>*</strong></font>Lưu ý: Quản trị 
            luôn có quyền với tất cả các chuyên mục</font></p></td>
        </tr>
      </table>
      </form></td>
  </tr>
  <tr> 
    <td height="35" align="center" valign="middle"><font size="2" face="Arial, Helvetica, sans-serif"><a href="javascript: window.close();">Đóng 
      cửa sổ</a></font></td>
  </tr>
</table>
</body>
</html>