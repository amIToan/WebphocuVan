<%session.CodePage=65001%>
<%Call PhanQuyen("QLyCongViec")%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<!--#include virtual="/include/func_calculate.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<%

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
  <table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td ></td>
    </tr>
    <tr>
      <td >
	  <form action="updatePhieu.asp" target="_blank" name="PhieuThuChi" method="post">
	    <table width="500" border="0" align="center" cellpadding="2" cellspacing="2" class="CTxtContent" style="border:#CCCCCC solid 1">
          <tr>
            <td colspan="2" class="CTieuDe" align="center" style="border-bottom:#CCCCCC solid 1"><img src="../../images/icons/Money20Icon.jpg" width="48" height="48" align="absmiddle"> TIỀN LƯƠNG </td>
          </tr>
          <tr>
            <td width="259" class="CTieuDeNhoNho"><img src="../../images/icons/jobs.gif" width="48" height="48" align="absmiddle"> Mức lương.</td>
            <td width="239" class="CTieuDeNhoNho"> <img src="../../images/icons/icon-Money32x32.gif" width="34" height="40" align="absmiddle"> Phụ cấp.</td>
          </tr>
          <tr>
            <td class="CTieuDeNhoNho"><img src="../../images/icons/icon_reward_details.gif" width="47" height="30" align="absmiddle"> Khen thưởng &amp; Kỷ luật.</td>
            <td class="CTieuDeNhoNho"> <img src="../../images/icons/industry-icon-insurance-48x.gif" width="48" height="48" align="absmiddle">Bảo hiểm XH &amp; YT.</td>
          </tr>
          <tr>
            <td class="CTieuDeNhoNho"><img src="../../images/icons/money_iconw.gif" width="48" height="48" align="absmiddle"> Tổng tiền lương.</td>
            <td>&nbsp;</td>
          </tr>
        </table>
	  </form>
</td>
    </tr>
  <td ></td>
  </tr>
  </table>
  <%Call Footer()%>
</body>
</html>