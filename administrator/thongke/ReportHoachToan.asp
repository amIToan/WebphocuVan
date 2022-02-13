<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/include/func_calculate.asp" -->
<!--#include virtual="/administrator/inc/funcInProduct.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<%
f_permission = administrator(false,session("user"),"m_report")
if f_permission < 2 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
	Ngay1=1
	Thang1=Month(now())-1
	Nam1=Year(now())
	Ngay2=Day(now())
	Thang2=Month(now())
	Nam2=Year(now())

%>

<html >
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=PAGE_TITLE%></title>
<link href="../../css/styles.css" rel="stylesheet" type="text/css"></head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<p>
	  <%
	  	Title_This_Page="Thống kê -> Hoách toán"
		Call header()
		Call Menu()
		
		
		%>
	</p>
	<div align="center"><strong>HOẠCH TOÁN HÀNG HÓA</strong></div>
	<form id="frmReportMH" name="frmReportMH" method="post" action="ReturnHoachToanDetail.asp">
	  <table width="590px" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td  background="../../images/T1.jpg" height="20"></td>
        </tr>
        <tr>
          <td background="../../images/t2.jpg">
		  <table width="95%" border="0" align="center" cellpadding="1" cellspacing="1">
            <tr>
              <td width="28%" class="CTxtContent"><input name="ctXuatKho" type="checkbox" id="ctXuatKho" value="1">
Chi tiến xuất kho<br>
  <input name="ctNhapKho" type="checkbox" id="ctNhapKho" value="1">
Chi tiết nhâp kho</td>
              <td width="33%" class="CTxtContent"><input name="ctTonKho" type="checkbox" id="ctTonKho" value="1">
Chi tiết tồn kho<br>
<input name="ctxseo" type="checkbox" id="ctxseo" value="1">
Chi tiết xseo xuất </td>
              <td width="39%" height="29" class="CTxtContent"><input name="ctDHHuy" type="checkbox" id="ctDHHuy" value="1">
              Chi tiết đơn hàng hủy <br></td>
            </tr>
            <tr>
              <td height="29" colspan="3" align="center" bordercolor="#FFFFFF"class="CTxtContent">Sắp xếp:&nbsp; Tăng dần
                <input name="RaOderBy" type="radio" value="0" checked <%if iOrderBy =0 then Response.Write("checked") end if%>>
                <font size="2" face="Verdana, Arial, Helvetica, sans-serif">/Giảm dần
                    <input name="RaOderBy" type="radio" value="1" <%if iOrderBy =1 then Response.Write("checked") end if%>>
              </font></font></span></td>
            </tr>
            <tr>
              <td height="29" colspan="3" align="center" bordercolor="#FFFFFF">
			  Thời gian:
                  <%
					Call List_Date_WithName(Ngay1,"DD","Ngay1")
					Call List_Month_WithName(Thang1,"MM","Thang1")
					Call List_Year_WithName(Nam1,"YYYY",2004,"Nam1")
				%>
                  <img src="../images/right.jpg" width="9" height="9" align="absmiddle" />
                  <%
					Call List_Date_WithName(Ngay2,"DD","Ngay2")
					Call List_Month_WithName(Thang2,"MM","Thang2")
					Call List_Year_WithName(Nam2,"YYYY",2004,"Nam2")
				%>
                <img  align="absbottom" src="/administrator/images/search_bt.gif" width="27" height="22"  style="cursor:pointer;" onClick="TestDate();" /> </font>                </p></td>
            </tr>
          </table></td>
        </tr>
        <tr>
          <td background="../../images/T3.jpg" height="8"></td>
        </tr>
      </table>
</form>    

<%Call Footer()%>
    
</body>
</html>

<script language="javascript">
	function TestDate()
	{
		if (document.frmReportMH.Ngay1.value==0)
		{
			alert('Bạn phải chọn ngày');
			document.frmReportMH.Ngay1.focus();
			return false;
		}
		if (document.frmReportMH.Thang1.value==0)
		{
			alert("Bạn chưa chọn tháng!");
			document.frmReportMH.Thang1.focus();
			return false;
		}
		if (document.frmReportMH.Nam1.value==0)
		{
			alert("Bạn chưa chọn năm!");
			document.frmReportMH.Nam1.focus();
			return false;
		}
		if (document.frmReportMH.Ngay2.value==0)
		{
			alert("Bạn chưa chọn ngày!");
			document.frmReportMH.Ngay2.focus();
			return false;
		}
		if (document.frmReportMH.Thang2.value==0)
		{
			alert("Bạn chưa chọn tháng!");
			document.frmReportMH.Thang2.focus();
			return false;
		}
		if (document.frmReportMH.Nam2.value==0)
		{
			alert("Bạn chưa chọn năm!");
			document.frmReportMH.Nam2.focus();
			return false;
		}
		document.frmReportMH.submit();
		}
		
	</script>
