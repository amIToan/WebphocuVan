<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<!--#include virtual="/administrator/inc/funcInProduct.asp" -->
<%
f_permission = administrator(false,session("user"),"m_report")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
 <%	
	inProductID = GetNumeric(Request.querystring("inProductID"),0)
 %>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
	<link href="../../css/styles.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table  width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
    <tr>
      <td width="48%"><div align="center"><img src="../../images/logoxseo128.png" width="157" height="57"></div></td>
      	<td width="52%"><div align="center"><em>www.xseo.com</em><br>
   	        <em>ĐT: <%=soDT%>  - Email: info@xseo.com</em></div></td>
    </tr>
    <tr>
      <td><div align="center"><strong><%=TenGD%></strong></div></td>
      <td width="53%"><div align="center"></div></td>
    </tr>
    <tr>
      <td><div align="center"><em>ĐC: <%=dcVanPhong%> </em></div></td>
      <td></td>
    </tr>
</table>
<br><br>
 <div align="center" class="author">
  <div align="center"><strong>HÓA ĐƠN NHẬP HÀNG</strong> </div>
</div>
<br><br>
<table width="590px" border="0" cellpadding="0" cellspacing="0" class="CTxtContent">
          <tr>
            <td  background="../../images/T1.jpg" height="20"></td>
          </tr>
          <tr>
            <td background="../../images/t2.jpg">
<%
		sql = "SELECT  inProductID,Maso,WorkerMuaHangID,WorkerThanhToanID,DateTime,Provider.ProviderName FROM inputProduct INNER JOIN Provider ON inputProduct.ProviderID =Provider.ProviderID where inProductID = '"& inProductID &"'"		
		set rs=Server.CreateObject("ADODB.Recordset")
		rs.open sql,Con,1 
if not rs.eof then
			Maso				=	rs("Maso")
			WorkerMuaHangID		=	rs("WorkerMuaHangID")
			WorkerThanhToanID	=	rs("WorkerThanhToanID")
			DateTime			=	rs("DateTime")
			ProviderName		=	rs("ProviderName")	
			
%></td>
          </tr>
          <tr>
            <td background="../../images/t2.jpg"><table width="95%"border="0"  align="center" cellpadding="1" cellspacing="1" class="CTxtContent">
                <tr>
                  <td >Mã số: <%=Maso%></td>
                </tr>
                <tr>
                  <td >Nhà cung cấp: <%=ProviderName%> </td>
                </tr>
                <tr>
                  <td > Ngày nhập:<%=day(DateTime)%>/<%=Month(DateTime)%>/<%=year(DateTime)%></td>
                </tr>
                <tr>
                  <td  >Nhân viên mua hàng: <b><%=GetNameNV(WorkerMuaHangID)%></b></td>
                </tr>
                <tr>
                  <td  >Thanh Toán : <b><%=GetNameNV(WorkerThanhToanID)%></b></td>
                </tr>

            </table>
<%end if%>			
			</td>
          </tr>
          <tr>
            <td background="../../images/T3.jpg" height="8"></td>
          </tr>
</table>
		<br>
	  
	  <table width="100%"  align="center" cellpadding="1"  cellspacing="1" border="0">
		  <tr>
			<td width="6%" align="center" style="<%=setStyleBorder(1,1,1,1)%>">STT</td>
			<td width="11%" align="center" style="<%=setStyleBorder(0,1,1,1)%>">Mã SP </td>
			<td width="47%" align="center" style="<%=setStyleBorder(0,1,1,1)%>">Tên SP </td>
			<td width="8%" align="center" style="<%=setStyleBorder(0,1,1,1)%>">Giá bìa </td>
			<td width="3%" align="center" style="<%=setStyleBorder(0,1,1,1)%>">VAT</td>
			<td width="3%" align="center" style="<%=setStyleBorder(0,1,1,1)%>">Nhập/Tồn </td>
			<td width="8%" align="center" style="<%=setStyleBorder(0,1,1,1)%>">Đơn giá</td>
			<td width="4%" align="center" style="<%=setStyleBorder(0,1,1,1)%>">CK<span class="style54">(%)</span></td>
			<td width="10%" align="center" style="<%=setStyleBorder(0,1,1,1)%>">T.tiền </td>
	    </tr>
	<%
	stt= 1
	TTien = 0
	sql = "SELECT Product.NewsID as NewsID, idsanpham,Title,tacgia,Giabia,Unit,Number,Price,VAT FROM Product INNER JOIN SanPhamNhap ON Product.NewsID = SanPhamNhap.NewsID WHERE (Product.inProductID = '"& inProductID &"') ORDER BY Product.ProductID DESC"
	set rsProduct=Server.CreateObject("ADODB.Recordset")
	rsProduct.open sql,Con,1 
	Do while not rsProduct.eof
	%>
		<tr>
		  <td align="center" style="<%=setStyleBorder(1,1,0,1)%>"><%=stt%></td>
			<td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=rsProduct("idsanpham")%></td>
			<td align="left" style="<%=setStyleBorder(0,1,0,1)%>">
				<%=LCase(rsProduct("Title"))%><br>
			<font class="CSubTitle">
				<i>Tác giả</i>: <%=rsProduct("tacgia")%>
			</font>	
		  </td>
			<td align="right" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(rsProduct("Giabia"))%></td>
		<%
		iNum 	=	rsProduct("Number")
		iPrice 	=	CDbl(rsProduct("Price"))
		Giabia 	= 	Clng(rsProduct("giabia"))
		iCK 	= 	100*(Giabia-iPrice)/Giabia
		iCK 	= 	round(iCK)
		iVAT 	=	GetNumeric(rsProduct("VAT"),0)
		Tien = iPrice*iNum
		Tien = Tien + Tien*iVAT/100
		TTien = TTien + Tien		

		%>

	<td align="center" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;<%=iVAT%></td>
	<td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=iNum%>/<%=GetNumInventoryGoods(rsProduct("NewsID"))%>
</td>
	<td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(iPrice)%></td>
	<td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=iCK%></td>
	<td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(LamTronTien(Tien))%></td>
		</tr>
	<%
		stt = stt +1
		rsProduct.movenext
	loop
	rsProduct.close
	%>
	<tr>
		  <td colspan="6" align="right" class="CFontVerdana10">Tổng:</td>
			<td colspan="3" align="center"><%=Dis_str_money(LamTronTien(TTien))%> đ</td>
  </tr>
</table>
		
		<br>
		
</body>
</html>
	
