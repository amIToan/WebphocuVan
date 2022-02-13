<%session.CodePage=65001%>
<%Call PhanQuyen("QLyCongViec")%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<!--#include virtual="/include/func_calculate.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->

<html>
<link href="../../css/styles.css" rel="stylesheet" type="text/css">
<body>
<%
function  fLapLichBieu()
%>
<table width="600" border="0" align="center" cellpadding="1" cellspacing="1" class="CTxtContent" style="border:#CCCCCC solid 1px;">
   <tr>
     <td colspan="2" class="CTieuDe" align="center"><a href="qlNhanSu.asp?isDetail=1&iCongViec=1" class="CFontVerdana10"><img src="../../images/icons/calendar_icon.gif" width="50" height="50" border="0" align="absmiddle"></a> LẬP LỊCH BIỂU </td>
   </tr>
   <tr>
     <td colspan="2">&nbsp;</td>
   </tr>
   <tr>
     <td width="128" align="right">&nbsp;</td>
     <td width="470" class="CTieuDeNho">Công việc loại A </td>
   </tr>
   <tr>
     <td>&nbsp;</td>
     <td>&nbsp;</td>
   </tr>
   <tr>
     <td>&nbsp;</td>
     <td class="CTieuDeNho">Công việc loại B </td>
   </tr>
   <tr>
     <td>&nbsp;</td>
     <td>&nbsp;</td>
   </tr>
   <tr>
     <td>&nbsp;</td>
     <td class="CTieuDeNho">Công việc loại C </td>
   </tr>
   <tr>
     <td>&nbsp;</td>
     <td>&nbsp;</td>
   </tr>
   <tr>
     <td>&nbsp;</td>
     <td class="CTieuDeNho">Công việc loại D </td>
   </tr>
   <tr>
     <td>&nbsp;</td>
     <td>&nbsp;</td>
   </tr>
   <tr>
     <td>&nbsp;</td>
     <td class="CTieuDeNho">Công việc loại E </td>
   </tr>
   <tr>
     <td>&nbsp;</td>
     <td>&nbsp;</td>
   </tr>
   <tr>
     <td>&nbsp;</td>
     <td class="CTieuDeNho">Công việc loại F </td>
   </tr>
   <tr>
     <td>&nbsp;</td>
     <td>&nbsp;</td>
   </tr>
   <tr>
     <td>&nbsp;</td>
     <td>&nbsp;</td>
   </tr>
 </table>
<%
end function
%>
</body>
</html>