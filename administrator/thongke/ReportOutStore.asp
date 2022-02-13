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
	Ngay1=Day(now())
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
	  	Title_This_Page="Thông kê-> Xuất kho"
		Call header()
		Call Menu()
		
		%>
	</p>
	<div  class="CTieuDe" align="center">THỐNG KÊ XUẤT HÀNG</div>
	<form id="frmReportMH" name="frmReportMH" method="post" action="ReturnBaoCaoXK.asp">
	  <table width="590px" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td  background="../../images/T1.jpg" height="20"></td>
        </tr>
        <tr>
          <td background="../../images/t2.jpg">
		  <table width="95%" border="0" align="center" cellpadding="1" cellspacing="1">
            
            <tr>
              <td width="50%" height="29" align="left" class="CTxtContent">Kiểm soát viên:
			    <%
					KSV=GetNumeric(Request.form("selKS"),0)
					call SelectNhanVien("selKS",KSV,1,0,0)
				%></td>
              <td width="50%" height="29" bordercolor="#FFFFFF" class="CTxtContent">Giao hàng :
                <%
						NMH=Clng(Request.Form("selMH"))
						call SelectNhanVien("selMH",NMH,1,"","")
						%></td>
            </tr>
            <tr>
              <td height="29" colspan="2" align="center" bordercolor="#FFFFFF"class="CTxtContent">
			Thu tiền:
			<%
			NVThutienID=Clng(Request.Form("selNVThutien"))
			call SelectNhanVien("selNVThutien",NVThutienID,1,0,0)
			%>			</td>
            </tr>
            <tr>
              <td height="29" colspan="2" align="center" bordercolor="#FFFFFF"class="CTxtContent">
			  <input name="txtMaOrTensach" type="text" id="txtMaOrTensach" value="<%=strMaorTenSach%>">  
                <select name="selMaorTenSach" id="selMaorTenSach">
                  <option value="0" <%if iMaorTenSach = 0 then%>selected<%end if%>></option>
                  <option value="1">Mã đơn hàng</option>
                  <option value="3">Tên khách</option>
                  <option value="4">Email</option>
                  <option value="5">Tel</option>
                  <option value="6">Mã sách</option>
				  <option value="8">Tiêu đề sách</option>
				  <option value="7">Địa chỉ</option>
                </select>              </td>
            </tr>
            <tr>
              <td height="29" colspan="2" align="center" bordercolor="#FFFFFF"class="CTxtContent">Sắp xếp:&nbsp; Tăng dần
                    <input name="RaOderBy" type="radio" value="0" checked <%if iOrderBy =0 then Response.Write("checked") end if%>>
                    <font size="2" face="Verdana, Arial, Helvetica, sans-serif">/Giảm dần
                    <input name="RaOderBy" type="radio" value="1" <%if iOrderBy =1 then Response.Write("checked") end if%>>
                    </font></font></span><strong>
                    <select name="selSearch">
                      <option value="SanPhamUser.SanPhamUser_ID" selected>Mã số</option>
                      <option value="SanPhamUser.SanPhamUser_Date">Theo ngày</option>
                      <option value="SanPhamUser.SanPhamUser_Name">Tên khách</option>
                    </select>
                    </strong></td>
            </tr>
            <tr>
              <td height="29" colspan="2" align="center" bordercolor="#FFFFFF"><span class="CTxtContent">Chi tiết:
                  <input name="iDetail" type="checkbox" id="iDetail" value="1">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Biểu đồ: 
              <input name="iBieuDo" type="checkbox" id="iBieuDo" value="1">
              </span></td>
            </tr>
            <tr>
              <td height="29" colspan="2" align="center" bordercolor="#FFFFFF">Thời gian:
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
              <img  align="absbottom" src="/administrator/images/search_bt.gif" width="27" height="22"  style="cursor:pointer;" onClick="TestDate();" /> </font>                  </p></td>
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
