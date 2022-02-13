<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/funcInProduct.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<%
f_permission = administrator(false,session("user"),"m_report")
if f_permission = 0 then
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
	  <%
	  	img	=	"../../images/icons/Home-icon.png"
		Title_This_Page="Thống kê -> Tình trạng kho."
		Call header()
		Call Menu()
		
		
		%>
	</p>
	<div align="center" class="CTieuDe">
	 THỐNG KÊ TRẢ KHO	</div>
	<form id="frmReportMH" name="frmReportMH" method="post" action="ReturnReStoreGoods.asp">
	  <table width="590px" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
        <tr>
          <td  background="../../images/T1.jpg" height="30" class="CTieuDeNhoNho">&nbsp;&nbsp;&nbsp;Lựa chọn thống kê theo yêu cầu</td>
        </tr>
        <tr>
          <td background="../../images/t2.jpg">
		  <table align="center" border="0" cellpadding="0" cellspacing="0" width="99%">
		  
		  <tr>
		  	<td width="75%" valign="top" style="border-right:#CCCCCC solid 1px;">
		  <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">		  
            <tr>
              <td height="29" colspan="2" align="center" bordercolor="#FFFFFF"class="CTxtContent">
			  			  
			  Tìm:		
		        <select name="selMaorTenSach" id="selMaorTenSach" onChange="javascript:onChangeOrder()">
		          <option value="0">Tất cả hóa đơn vào</option>
		  <option value="1" >Hóa đơn nhập</option>
		  <option value="2" >Nhà cung cấp</option>
		  <option value="10">------------------</option>
		  <option value="3">Tất cả đầu sách</option>
		  <option value="4">Mã sách</option>
		  <option value="5">Tên sách</option>
		  <option value="10">------------------</option>
		                        </select>
		&nbsp;&nbsp;&nbsp;&nbsp;
		<%	if iMaorTenSach =  4 then 
				DisConSearch2	=	""
				DisConSearch1	=	"none"
			else
				DisConSearch2	="none"
				DisConSearch1=""
			end if
			
			%>		
			
			<span id="ConSearch1" style="display:<%=DisConSearch1%>">
	      <input name="txtMaOrTensach" type="text" id="txtMaOrTensach" value="<%=strMaorTenSach%>"> 
		</span>
				<span id="ConSearch2" style="display:<%=DisConSearch2%>">
		<%call SelectProvider("selProvider",ProviderID)%> 
		</span>		</td>
            </tr>

			
            <tr>
              <td height="29" colspan="2" align="center" bordercolor="#FFFFFF" class="CTxtContent">
			  <span id="SelDatetime" style="display:">
			  Ngày nhập:
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
               </span>			  </td>
            </tr>
            <tr>
              <td height="29" colspan="2" align="center" bordercolor="#FFFFFF" class="CTxtContent">
			  <span id="ShowInventory" style="display:">
			  <input name="iInventory" type="checkbox" id="iInventory" value="1">
                Chỉ xem Hóa đơn có trả kho  &nbsp; &nbsp;&nbsp;
                  <input name="iBookInventory" type="checkbox" id="iBookInventory" value="1">
                  Chỉ xem sách trả 
				  </span>
				  <span id="ShowDetailInventore" style="display:none">
				  Xem cơ bản
				  <input name="iDetailInventore" type="radio" id="iDetailInventore" value="0" checked > 
				  /
				  Xem chi tiết nhập kho
				  	<input name="iDetailInventore" type="radio" id="iDetailInventore" value="1"> / Xem ngày nhập <input name="iDetailInventore" type="radio" id="iDetailInventore" value="2">
				  </span>
				  </td>
            </tr>

            <tr>
              <td height="29" colspan="2" align="center" bordercolor="#FFFFFF"class="CTxtContent">
		<span id="ConSearch3" style="display:">
		Sắp xếp
		<input name="SelSort" type="text" class="CTextBox_NoBorder" id="SelSort" value="ngày nhập" size="12">
		 Mặc định
		 <input name="RaOderBy" type="radio" value="3" <%if iOrderBy =1 then Response.Write("checked") end if%>>
		 / Tăng dần
                    <input name="RaOderBy" type="radio" value="1" <%if iOrderBy =0 then Response.Write("checked") end if%>>
                    / Giảm dần
                    <input name="RaOderBy" type="radio" value="2" <%if iOrderBy =1 then Response.Write("checked") end if%>>	
		</span></td>
            </tr>			
            <tr>
              <td height="29" colspan="2" align="center" bordercolor="#FFFFFF" class="CTxtContent"> 
			  <input type="button" name="Submit" value=" Thực hiện "  onClick="javascript: TestDate();" ></td>
            </tr>		
          </table>			
			</td>
			<td valign="top"><table width="100%" border="0" cellspacing="2" cellpadding="2">

              <tr>
                <td><a href="ReportInputStore.asp"><img src="../../images/icons/InputStoreIcon.png" width="48" height="48" align="absmiddle" border="0"> Nhập kho</a> </td>
              </tr>
              <tr>
                <td><a href="ReportInputStore.asp"><img src="../../images/icons/OutStoreIcon.png" width="48" height="48" align="absmiddle" border="0"></a><a href="ReportOutStore.asp"> Xuất kho </a></td>
              </tr>			  
              <tr>
                <td><a href="ReportStoreGoods.asp"><img src="../../images/icons/Home-icon.png" width="48" height="48" border="0" align="absmiddle"> Tồn kho </a> </td>
              </tr>
            </table></td>
		  </tr>
		  </table>
		  

		  </td>
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
<script language="javascript">
 function onChangeOrder()
 {
 	if(document.frmReportMH.selMaorTenSach.options[2].selected==true)
	{
		document.getElementById("ConSearch1").style.display="none";
		document.getElementById("ConSearch2").style.display="";
		document.getElementById("ConSearch3").style.display="none";
	}
	else
	{
		document.getElementById("ConSearch1").style.display="";
		document.getElementById("ConSearch2").style.display="none";
		
	}
	

	if(document.frmReportMH.selMaorTenSach.options[1].selected==true)
	{
		document.getElementById("SelDatetime").style.display="none";
		document.getElementById("ShowInventory").style.display="";
		document.getElementById("ConSearch3").style.display="none";
		}
	else{
		document.getElementById("SelDatetime").style.display="";
		document.getElementById("ShowInventory").style.display="none";
		}
		

	if((document.frmReportMH.selMaorTenSach.options[4].selected==true)||
		(document.frmReportMH.selMaorTenSach.options[5].selected==true)||
		(document.frmReportMH.selMaorTenSach.options[6].selected==true))
	{		
		document.getElementById("SelDatetime").style.display="none";
		document.getElementById("ShowInventory").style.display="none";
		document.getElementById("ConSearch3").style.display="none";
		document.getElementById("ShowDetailInventore").style.display="";
	}
	else
	{
		document.getElementById("SelDatetime").style.display="";
		document.getElementById("ShowInventory").style.display="";
		document.getElementById("ShowDetailInventore").style.display="none";
	}
	if(document.frmReportMH.selMaorTenSach.options[0].selected==true)
	{
		document.frmReportMH.SelSort.value	=	'ngày nhập'
		document.getElementById("ConSearch3").style.display="";
		document.getElementById("ConSearch1").style.display="none";
		}

		
	if(document.frmReportMH.selMaorTenSach.options[4].selected==true)
	{
		document.frmReportMH.SelSort.value	=	'tiêu đề'
		document.getElementById("ConSearch3").style.display="";
		document.getElementById("ConSearch1").style.display="none";
		}	
	
 }
</script>