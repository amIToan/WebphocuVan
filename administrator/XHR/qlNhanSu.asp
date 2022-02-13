<%session.CodePage=65001%>
<%Call PhanQuyen("QLyCongViec")%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<!--#include virtual="/include/func_calculate.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<!--#include virtual="/administrator/nhansu/fun_CongViec.asp" -->
<%
	isDetail	= GetNumeric(Request.QueryString("isDetail"),0)
	iCongViec	= GetNumeric(Request.QueryString("iCongViec"),0)
%>
<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
<LINK href="/administrator/inc/admstyle.css" type=text/css rel=stylesheet>
<link href="../../css/styles.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style57 {color: #FF0000}
a:link {
	text-decoration: none;
}
a:visited {
	text-decoration: none;
}
a:hover {
	text-decoration: none;
}
a:active {
	text-decoration: none;
}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
  <%
  	img	=	"../../images/icons/Meber.gif"
	Title_This_Page="Quản trị nhân sự-> Quản lý công việc."
	Call header()
	Call Menu()
%>
<%if isDetail = 0 then%>
  <table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td ></td>
    </tr>
    <tr>
      <td >
	  
	    <table width="500" border="0" align="center" cellpadding="2" cellspacing="2" class="CTxtContent" style="border:#CCCCCC solid 1">
          <tr>
            <td colspan="2" class="CTieuDe" align="center" style="border-bottom:#CCCCCC solid 1"><span class="CTieuDeNhoNho"><img src="../../images/icons/AlarmClocl20Icon.jpg" width="48" height="48" align="absmiddle"></span>QUẢN TRỊ CÔNG VIỆC </td>
          </tr>
          <tr>
            <td width="259" class="CTieuDeNhoNho"> 
			<a href="qlNhanSu.asp?isDetail=1&iCongViec=1" class="CFontVerdana10"><img src="../../images/icons/calendar_icon.gif" width="50" height="50" border="0" align="absmiddle"></a> Lập kế hoạch công việc						</td>
            <td width="239" class="CTieuDeNhoNho"> <a href="qlNhanSu.asp?isDetail=1&iCongViec=1" class="CFontVerdana10"><img src="../../images/icons/Aqua-Smooth-Folder-Works-ic.gif" alt=" " width="48" height="48" border="0" align="absmiddle"></a>Theo dõi công việc</td>
          </tr>
          <tr>
            <td class="CTieuDeNhoNho"><img src="../../images/icons/reports_icon.gif" width="50" height="45" align="absmiddle"> Báo cáo kết quả </td>
            <td>&nbsp;</td>
          </tr>
        </table>

</td>
    </tr>
  <td ></td>
  </tr>
  </table>
  <%end if%>
 <%
	Select 	case iCongViec
			case 1
				fLapLichBieu()
			case 2
			case 3
	end select
 
 %>
  <%Call Footer()%>
</body>
</html>