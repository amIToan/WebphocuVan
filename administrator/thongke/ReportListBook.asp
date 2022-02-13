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
	  <%Title_This_Page="Thống kê nhập hàng"
		Call header()
		Call Menu()
		
		
		%>
	</p>
	
	<form id="frmReportMH" name="frmReportMH" method="post" action="ReturnListBook.asp">
	  <table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td ></td>
        </tr>
        <tr>
          <td >
		  <table align="center" cellpadding="2" cellspacing="2" width="770">
	<tr>
	  <td align="center" valign="middle" class="CTxtContent" > <strong class="author">THỐNG KÊ DANH SÁCH SÁCH </strong></td>
	  </tr>
	<tr> 
		<td align="center" valign="middle" width="100%" class="CTxtContent">
Ngày:&nbsp;			
<%
			Call List_Date_WithName(Ngay1,"DD","Ngay1")
			Call List_Month_WithName(Thang1,"MM","Thang1")
			Call  List_Year_WithName(Nam1,"YYYY",2004,"Nam1")
%>
 <img src="../images/right.jpg" width="9" height="9" align="absmiddle"> 		
<%
			Call List_Date_WithName(Ngay2,"DD","Ngay2")
			Call List_Month_WithName(Thang2,"MM","Thang2")
			Call  List_Year_WithName(Nam2,"YYYY",2004,"Nam2")
%>    	</td>
	</tr>
	<tr>
	  <td align="center" valign="middle" class="CTxtContent">
	   
<input name="keyword" type="text" value="<%=Replace(KeyWord,"""","&quot;")%>" size="25"  onKeyUp="initTyper(this);">&nbsp;&nbsp;&nbsp;&nbsp;
                <select name="select_filter" id="select_filter">
                  <option value="1" <% if seach_filter = 1  then Response.Write("selected")%>>Tìm theo tiêu trí</option>
                  <option value="2" <% if seach_filter = 2  then Response.Write("selected")%>>Tiêu đề</option>
				  <option value="10" <% if seach_filter = 10  then Response.Write("selected")%>>Loại sách</option>
                  <option value="5" <% if seach_filter = 5  then Response.Write("selected")%>>Tác giả</option>				  
                  <option value="4" <% if seach_filter = 4  then Response.Write("selected")%>>Nhà xuất bản</option>				  
				  <option value="9" <% if seach_filter = 9  then Response.Write("selected")%>>Loại sách</option>
				  <option value="8" <% if seach_filter = 8  then Response.Write("selected")%>>Mã sách</option>
                  <option value="3" <% if seach_filter = 3  then Response.Write("selected")%>>Nội dung</option>
                  <option value="6" <% if seach_filter = 6  then Response.Write("selected")%>>Giá bìa</option>
				  <option value="7" <% if seach_filter = 7  then Response.Write("selected")%>>Tất cả</option>	
				  <option value="11" <% if seach_filter = 11  then Response.Write("selected")%>>Nhà cung cấp</option>			  
                </select><br>
                <span class="CTxtContent">
                <input name="isSapXep" type="radio" value="1" <% if isSapXep = 1  then Response.Write("checked")%>/>
Giảm dần
<input name="isSapXep" type="radio" value="2" <% if isSapXep = 2  then Response.Write("checked")%>/>
Tăng dần</span>
                <input type="button" name="Button" value="  OK  " onClick="javascript:TestDate()">
			<input type="hidden" name="action" value="Search">	  </td>
    </tr>
</table>
		  </td>
        </tr>
  <td ></td>
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
