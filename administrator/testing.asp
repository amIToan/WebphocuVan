<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%  
'	XSOFT
'  	(C) Copyright XSOFT Corp. 2007
'  	**************************
' 	Cong ty tu van thiet ke va phat trien phan mem  
'  	Quan ly nhan su, ban hang, ton kho, tai chinh ke toan, tai chinh gia dinh.
'  	Thiet ke website, thiet ke logo, catalog.
'  	website:www.xsoftvn.com
'  	email:info@xsoftvn.com – DT:04.2922.446
%>
<%
'********************************************************************/
' file name      	: index.asp
' Create /Modifed by: tuannv
' Description		: file giao dien trang chu.
'********************************************************************/
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="verify-v1" content="4MCM8uFP8ATc+yjOGsgtW58O5uSre/q1Eqa2q/sCP8A=" />
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/include/func_common.asp" -->
<!--#include virtual="/include/func_DateTime.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<html>

<head>
<title>Nhà sách trên mạng, sách giảm giá, Bán sách ọnine,  mua sách online, </title>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<meta name="robots" content="INDEX,FOLLOW" />
<meta http-equiv="REFRESH" content="5400" />	
<meta name="description" content="<%=description_meta%>">
<meta name="keywords" content="<%=description_meta%>">
<meta http-equiv="Refresh" Content="1800">
<meta http-equiv="Expires" Content="600">
	
<link rel="stylesheet" type="text/css" href="css/styles.css">
<link href="css/styles.css" rel="stylesheet" type="text/css">
<link href="images/Logo/Xlogo.ico" rel="icon" type="image/x-icon" />
<link href="images/Logo/Xlogo.ico" rel="shortcut icon" />	
<SCRIPT language=JavaScript1.2 src="include/news.js"></SCRIPT>
<SCRIPT language=JavaScript1.2 src="include/vietuni.js"></SCRIPT>
<script language="javascript" type="text/javascript" src="/include/Scripts/jquery.js"></script>
<script language="javascript" type="text/javascript" src="/include/Scripts/jquery.easing.js"></script>
<script language="javascript" type="text/javascript" src="/include/Scripts/script.js"></script>
   <script src="Scripts/jquery-1.8.2.min.js"></script>
<link href="css/menu.css" rel="stylesheet" />

</head>
<body leftMargin="0" topMargin="0" marginwidth="0" marginheight="0">
 <%
	if GetNumeric(Request.QueryString("IDTB"),0)<> 0 then
		img	="../images/icons/announce_icon_white.gif"
		Title_This_Page="Bản thông báo"	
	else
		img	="../images/icons/key-48x48.gif"
		Title_This_Page="Đăng nhập thành công"
	end if	
	Call header()
	Call menu()
    call MenuVertical()
%>
    <table align="center">
        <tr>
            <td>
               <img src="../images/testing.jpg"  />
            </td>
        </tr>
    </table>

</body>
</html>

<script type="text/javascript">

