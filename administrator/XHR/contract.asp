<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<!--#include virtual="/include/func_calculate.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<%
f_permission = administrator(false,session("user"),"m_human")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
iStatus		=	Request.QueryString("iStatus")
NhanvienID	=	Request.QueryString("NhanvienID")	
if iStatus	=	"edit" or iStatus="view" then
	ContractID = Request.QueryString("contractid")
	sql = "SELECT  * FROM StaffContract where ID =" & ContractID
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open sql,Con,3
	If not rs.eof Then
		NhanVienID	=	rs("NhanVienID")
		Congviec	=	Trim(rs("Congviec"))
		'Congviec	=	Replace(Congviec,"<br>",chr(13) & chr(10))
		NgayHD		=	rs("NgayHD")
		EndDate		=	rs("EndDate")	
		ChucdanhID	=	rs("ChucdanhID")
		PhongID		=	rs("PhongID")
		isMember	=	rs("isMember")
        moneyoff    =   rs("moneyoff")
		luong		=	GetNumeric(rs("luong"),0)
		Heso		=	GetNumeric(rs("Heso"),0)
		trachnhiem	=	GetNumeric(rs("trachnhiem"),0)
		luong_BH	=	GetNumeric(rs("luong_BH"),0)
		Phucap		=	GetNumeric(rs("Phucap"),0)
		ThuTruongID	=	rs("ThuTruongID")
		Dongy		=	rs("Dongy")
		Kynhan		=	rs("Kynhan")
		Dieukhoan	=	TRIM(rs("Dieukhoan"))
		'Dieukhoan	=	Replace(Dieukhoan,"<br>",chr(13) & chr(10))
		set rs=nothing
	end if	
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>HỢP ĐỒNG NHÂN SỰ</title>
<script src='../inc/news.js'></script>
<link href="../../css/styles.css" rel="stylesheet" type="text/css">
    <style type="text/css">
        .auto-style1 {
            font-family: Arial;
            font-size: 11pt;
            color: #161616;
            line-height: 180%;
            letter-spacing: 1px;
            padding: 2px;
            margin: auto;
            width: 254px;
        }
        .auto-style2 {
            width: 254px;
        }
    </style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
  <%
	img	="../../images/icons/man20Icon.jpg"
	Title_This_Page="HỢP ĐỒNG LAO ĐỘNG  -> Cán bộ công ty"

	Call header()
	Call Menu()
%>
  <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" style="border:#CCCCCC solid 1px;" class="CTxtContent">
  <form name="fStaff" action="upStaff.asp" method="post" enctype="multipart/form-data">
    <tr>
      <td colspan="4" background="../../images/TabChinh.gif" height="29" style="background-repeat:no-repeat" class="CTieuDeNho">&nbsp;HỢP ĐỒNG LAO ĐỘNG </td>
    </tr>
	<%
	sql = "SELECT  * FROM NhanVien where NhanVienID ='"& NhanvienID &"'"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open sql,Con,3
	If not rs.eof Then
		Ho_Ten	=	rs("Ho_Ten")
		CMT		=	rs("CMT")
		imgNV	=	rs("imgNV")
		eChuKy	=	rs("eChuKy")
		set rs=nothing
	end if					
	%>
    <tr>
      <td align="right"  class="auto-style1">Họ và Tên:</td>
      <td width="477"class="CTxtContent" ><b><%=Ho_Ten%></b></td>
      <td colspan="2" rowspan="10"  class="CTxtContent" align="center" valign="top" style="border:#3399FF solid 1px;"> 
	    <font class="CTieuDeNhoNho">ẢNH CÁN BỘ</font> <br>
        <%if imgNV<>"" then%>
        <img src="<%=imgNV%>" height="180" border="0">
        <%end if%>		</td>
    </tr>
    <tr>
      <td   align="right" class="auto-style2">CMT:</td>
      <td class="CTxtContent" ><%=CMT%></td>
    </tr>
    <tr>
      <td  align="right" class="auto-style2">Hợp đồng từ ngày:</td>
      <td class="CTxtContent">
        <select name="NgayHD" id="NgayHD">
          <option value="0" selected="selected">ngày</option>
          <%For i = 1 to 31 step 1%>
          <option value="<%=i%>" <%if i=day(NgayHD) then%> selected="selected" <%end if%>><%=i%></option>
          <%Next%>
        </select>
/
<select name="ThangHD" id="ThangHD">
  <option value="0" selected="selected">tháng</option>
  <%For i = 1 to 12 step 1%>
  <option value="<%=i%>" <%if i=Month(NgayHD) then%> selected="selected" <%end if%>><%=i%></option>
  <%Next%>
</select>
/
<select name="NamHD" id="NamHD">
  <option value="0" selected="selected">năm</option>
  <%For i = year(now())+10 to 1990 step -1%>
  <option value="<%=i%>" <%if i=Year(NgayHD) then%> selected="selected" <%end if%>><%=i%></option>
  <%Next%>
</select>     </td>
    </tr>
    <tr>
      <td class="auto-style1"  align="right">Đến ngày:</td>
      <td class="CTxtContent"><select name="NgayEnd" id="NgayEnd">
        <option value="0" selected="selected">ngày</option>
        <%For i = 1 to 31 step 1%>
        <option value="<%=i%>" <%if i=day(EndDate) then%> selected="selected" <%end if%>><%=i%></option>
        <%Next%>
      </select>
/
<select name="ThangEnd" id="ThangEnd">
  <option value="0" selected="selected">tháng</option>
  <%For i = 1 to 12 step 1%>
  <option value="<%=i%>" <%if i=Month(EndDate) then%> selected="selected" <%end if%>><%=i%></option>
  <%Next%>
</select>
/
<select name="NamEnd" id="NamEnd">
  <option value="0" selected="selected">năm</option>
  <%For i = year(now())+10 to 1980 step -1%>
  <option value="<%=i%>" <%if i=Year(EndDate) then%> selected="selected" <%end if%>><%=i%></option>
  <%Next%>
</select></td>
    </tr>
    
    
    <tr>
      	<td class="auto-style1"  align="right">Chức danh:</td>
      	<td class="CTxtContent"><%call SelectChucDanh(ChucdanhID,"ChucdanhID")%></td>
    </tr>
    <tr>
      <td class="auto-style1"  align="right">Phòng: </td>
      <td class="CTxtContent"><%call selectroom(PhongID,"PhongID")%></td>
    </tr>
    <tr>
      <td class="auto-style1"  align="right">Loại hợp đồng: </td>
      <td class="CTxtContent">  
	  	<select name="isMember">
		<option value="1" <%if isMember=1 then%> selected="selected"<%end if%>>Đang hoạt động</option>
		<option value="0" <%if isMember=0 then%> selected="selected"<%end if%>>Đã nghỉ việc</option>
		<option value="5" <%if isMember=5 then%> selected="selected"<%end if%>>Hết hợp đồng</option>
		<option value="2" <%if isMember=2 then%> selected="selected"<%end if%>>Bán thời gian</option>
		<option value="3" <%if isMember=3 then%> selected="selected"<%end if%>>Đối tác</option>
		<option value="4" <%if isMember=4 then%> selected="selected"<%end if%>>Khác</option>
  		</select>  </td>
    </tr>
    <tr>
      <td class="auto-style1"  align="right">Hưởng CK:</td>
      <td class="CTxtContent">  
        <input name="moneyoff" type="text" class="CTextBoxUnder" id="moneyoff" value="<%=Dis_str_money(moneyoff)%>" maxlength="2" onKeyUp="javascript: DisMoneyThis(this);" size="5">
          %<br /> <span class="CSubTitle">Là tỷ lệ % nếu khách hàng nhập mã bán hàng, hoặc số tiền được hưởng cho <%=Ho_Ten%>một sản phẩm bán ra của XSEO</span></td>
    </tr>
    <tr>
      <td class="auto-style1"  align="right">Giá bán</td>
      <td class="CTxtContent">  
          &nbsp;</td>
    </tr>
    <tr>
      <td class="auto-style1"  align="right">Mức lương: </td>
      <td class="CTxtContent">
        <input name="luong" type="text" class="CTextBoxUnder" id="luong" value="<%=Dis_str_money(luong)%>" maxlength="50" onKeyUp="javascript: DisMoneyThis(this);" size="15">
      </td>
    </tr>
    <tr>
      <td class="auto-style1"  align="right">Hệ số:</td>
      <td class="CFontVerdana10"><input name="Heso" type="text" class="CTextBoxUnder" id="Heso" value="<%=Heso%>" size="3" onBlur="checkIsNumber(this)"></td>
    </tr>
    
    <tr>
      <td class="auto-style1"  align="right">Lương trách nhiệm: </td>
      <td class="CTxtContent"><input name="trachnhiem" type="text" class="CTextBoxUnder" id="trachnhiem" value="<%=Dis_str_money(trachnhiem)%>" maxlength="50" onKeyUp="javascript: DisMoneyThis(this);" size="15"></td>
    </tr>
    <tr>
      <td class="auto-style1"  align="right">Phụ cấp: </td>
      <td class="CTxtContent"><input name="Phucap" type="text" class="CTextBoxUnder" id="Phucap" value="<%=Dis_str_money(Phucap)%>" maxlength="50" onKeyUp="javascript: DisMoneyThis(this);" size="15"></td>
      <td colspan="2"  class="CTxtContent" align="center" valign="top" >&nbsp;</td>
    </tr>
    <tr>
      <td  align="right" class="auto-style1">Lương bảo hiểm: </td>
      <td class="CTxtContent">
	  <input name="luong_BH" type="text" class="CTextBoxUnder" id="luong_BH" value="<%=Dis_str_money(luong_BH)%>" maxlength="50" onKeyUp="javascript: DisMoneyThis(this);" size="15">	  </td>
      <td width="190"  align="right" class="CTxtContent" >&nbsp;</td>
      <td width="350" class="CTxtContent"  >&nbsp;</td>
    </tr>
    <tr>
      <td colspan="4"  class="CTxtContent" valign="top" align="center" >Yêu cầu công việc : <br>
        <textarea name="Congviec" cols="75" rows="15" id="Congviec" ><%=Congviec%></textarea></td>
    </tr>
    <tr>
      <td colspan="4"  class="CTxtContent" valign="top" align="center">Các điều khoản:<br>
	    <textarea name="Dieukhoan" cols="75" rows="20" id="Dieukhoan" ><%=Dieukhoan%></textarea>	  </td>
    </tr>
    <tr>
      <td colspan="2" class="CTxtContent" align="center" valign="top">
	  <font class="CTieuDeNhoNho" >
	  KÝ NHẬN</font><br>
  
	  <%if Kynhan<>0 then
      	call GetEChuKy(NhanVienID,0,0,0)%>
          <input type="hidden" name="StaffSigna" value="1">
          <%else%>
          <input name="StaffSigna" type="checkbox" id="StaffSigna" value="1">
		  <%end if%>
			 <br>
		<font class="CTieuDeNhoNho">
		<b><%=UCASE(Ho_Ten)%></b></font>
		          <br>
          <%if iStatus	=	"edit" and f_permission > 2 and eChuKy<>"" then%>
          <input type="checkbox" name="RemoveImageeChuKy" value="1">
        	Hủy ký
        <%end if%>	  </td>
      <td colspan="2"  align="center" valign="top" class="CTxtContent"><font class="CTieuDeNhoNho" >ĐẠI DIỆN CÔNG TY</font><br>
		<%if dongy <> 0  then
		call GetEChuKy(ThuTruongID,0,0,0)
		Response.Write("<input name=""CDongY"" type=""hidden"" value=""1"">")
		elseif GetIDNhanVienUserName(session("user")) = ThuTruongID then
		%>
		Ký: 
		<input name="CDongY" type="checkbox" id="CDongY" value="1" >
		<%end if%>	
		<br>
		<%if dongy <> 0 then 
		Response.Write("<font class=""CTieuDeNhoNho""><b>"&UCASE(GetNameNV(ThuTruongID))&"</b></font>")
		%>
			<input name="ThuTruongID" type="hidden" value="<%=ThuTruongID%>">
		<% 			   
		else 
			call SelectNhanVien("ThuTruongID",ThuTruongID,1,"Ban giám đốc","phó giám đốc") 
		end if%>	  
		<br>
          <%if iStatus	=	"edit" and f_permission > 2 and eChuKy<>"" then%>
          <input type="checkbox" name="RemoveImageeChuKy0" value="1">
        	Hủy ký
        <%end if%>
		</td>
    </tr>
    <tr>
      <td colspan="4" align="center">
	  <%if iStatus<>"view" then%>
	  <input type="submit" value="    Cập nhật   " name="button" >
	  <%end if%>
      <input type="reset" name="Submit2" value="   Nhập lại   ">
	  <input type="hidden" name="iStatus" value="<%=iStatus%>">
	   <input type="hidden" name="Catup" value="contract">
	    <input type="hidden" name="ContractID" value="<%=ContractID%>">
	  <input type="hidden" name="NhanvienID" value="<%=NhanvienID%>">	  
	  </td>
    </tr>
	</form>	
    <tr>
      <td colspan="4" class="CTxtContent" height="100">&nbsp;</td>
    </tr>
  </table>

<script>VISUAL=4; FULLCTRL=1;</script>
<script src='../js/quickbuild.js'></script>
<script>changetoIframeEditor(document.forms[0].Congviec)</script>
<script>changetoIframeEditor(document.forms[0].Dieukhoan)</script>
</body>
</html>