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
iStatus	=	Request.QueryString("iStatus")
if iStatus	=	"edit" then
	StaffContractID = Request.QueryString("StaffContractID")
	sql = "SELECT  * FROM NhanVien where NhanVienID =" & StaffContractID
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open sql,Con,3
	If not rs.eof Then
		Ho_Ten	=	rs("Ho_Ten")
		CMT		=	rs("CMT")
        salecode=   rs("salecode")
		NgaySinh=	rs("NgaySinh")
		NgayCap	=	rs("NgayCap")	
		Noicap	=	rs("Noicap")
		Hocvan	=	rs("Hocvan")
		DanToc	=	rs("DanToc")
		Tel		=	rs("Tel")
		Mobile	=	rs("Mobile")
		Email	=	rs("Email")
		Cu_tru	=	Trim(rs("Cu_tru"))
		Diachi	=	Trim(rs("Diachi"))
		imgNV	=	rs("imgNV")
		
		infoBoMe=	Trim(rs("infoBoMe"))
'		infoBoMe=	Replace(infoBoMe,"<br>",chr(13) & chr(10))			
		infoAnhEm=	Trim(rs("infoAnhEm"))
'		infoAnhEm=	Replace(infoAnhEm,"<br>",chr(13) & chr(10))
		infoVoChongCon=Trim(rs("infoVoChongCon"))
'		infoVoChongCon=	Replace(infoVoChongCon,"<br>",chr(13) & chr(10))
		Hoatdongbanthan=Trim(rs("Hoatdongbanthan"))
'		Hoatdongbanthan=	Replace(Hoatdongbanthan,"<br>",chr(13) & chr(10))
		password_partner	=	rs("password_partner")
		
		TK_NganHang=rs("TK_NganHang")
		BankID	=	rs("BankID")
		
		eChuKy	=	rs("eChuKy")
		set rs=nothing
	end if	
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>xseo - Hồ sơ nhân sự</title>
<script src='../inc/news.js'></script>
<link href="../../css/styles.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
  <%
	img	="../../images/icons/man20Icon.jpg"
	Title_This_Page="Thông tin cá nhân  -> Cán bộ công ty"

	Call header()
	Call Menu()
%>
  <form name="fStaff" action="upStaff.asp?Catup=staff" method="post" enctype="multipart/form-data">
  <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" style="border:#CCCCCC solid 1px;">

    <tr>
      <td colspan="4" background="../../images/TabChinh.gif" height="29" style="background-repeat:no-repeat" class="CTieuDeNho">&nbsp;THÔNG TIN CÁN BỘ </td>
    </tr>
    <tr>
      <td width="177" align="right"  class="CTxtContent">Họ và Tên:</td>
      <td width="477" ><input name="Ho_Ten" type="text" size="35" value="<%=Ho_Ten%>" class="CTextBoxUnder"></td>
      <td colspan="2" rowspan="12"  class="CTxtContent" align="center" valign="top" style="border:#3399FF solid 1px;"> 
	    <font class="CTieuDeNhoNho">ẢNH CÁN BỘ</font> <br>
        <%if imgNV<>"" then%>
        <img src="<%=imgNV%>" height="180" border="0"><br>
		<input type="hidden" value="<%=imgNV%>" name="imgNV">
        <%end if%>
        <input name="imgNVfile" type="file" id="imgNVfile" size="17">
        <%if iStatus	=	"edit" and Session("iQuanTri") = 1 and imgNV<>"" then%>
        <input type="checkbox" name="RemoveImageimgNV" value="1">
Xóa  ảnh
<%end if%></td>
    </tr>

    <tr>
      <td class="CTxtContent" align="right">Ngày sinh: </td>
      <td class="CTxtContent">
        <select name="NgaySinh" id="NgaySinh">
          <option value="0">ngày</option>
          <%For i = 1 to 31 step 1%>
          <option value="<%=i%>" <%if i=day(NgaySinh) then%> selected="selected" <%end if%>><%=i%></option>
          <%Next%>
        </select>
/
<select name="ThangSinh" id="ThangSinh">
  <option value="0" >tháng</option>
  <%For i = 1 to 12 step 1%>
  <option value="<%=i%>" <%if i=Month(NgaySinh) then%> selected="selected" <%end if%>><%=i%></option>
  <%Next%>
</select>
/
<select name="NamSinh" id="NamSinh">
  <option value="0">năm</option>
  <%For i = year(now())-10 to year(now())-50 step -1%>
  <option value="<%=i%>" <%if i=Year(NgaySinh) then%> selected="selected" <%end if%>><%=i%></option>
  <%Next%>
</select>     </td>
    </tr>
    <tr>
      <td class="CTxtContent"  align="right">CMT:</td>
      <td class="CTxtContent" ><input name="CMT" type="text" size="25" value="<%=CMT%>" class="CTextBoxUnder"> </td>
    </tr>
    <tr>
      <td class="CTxtContent"  align="right">Mã bán hàng:</td>
      <td class="CTxtContent" ><input name="salecode" type="text" size="25" value="<%=salecode%>" class="CTextBoxUnder"></td>
    </tr>
    <tr>
      <td class="CTxtContent"  align="right">Password:
        </td>
      <td class="CTxtContent" >
        <input name="password_partner" type="text" class="CTextBoxUnder" id="password_partner" value="<%=password_partner%>" size="25"></td>
    </tr>
    <tr>
      <td class="CTxtContent"  align="right">Ngày cấp</td>
      <td class="CTxtContent"><select name="NgayCap" id="NgayCap">
        <option value="0">ngày</option>
        <%For i = 1 to 31 step 1%>
        <option value="<%=i%>" <%if i=day(NgayCap) then%> selected="selected" <%end if%>><%=i%></option>
        <%Next%>
      </select>
/
<select name="ThangCap" id="ThangCap">
  <option value="0">tháng</option>
  <%For i = 1 to 12 step 1%>
  <option value="<%=i%>" <%if i=Month(NgayCap) then%> selected="selected" <%end if%>><%=i%></option>
  <%Next%>
</select>
/
<select name="NamCap" id="NamCap">
  <option value="0">năm</option>
  <%For i = year(now()) to year(now())-20 step -1%>
  <option value="<%=i%>" <%if i=Year(NgayCap) then%> selected="selected" <%end if%>><%=i%></option>
  <%Next%>
</select></td>
    </tr>
    
    
    <tr>
      <td class="CTxtContent"  align="right">Nơi cấp:</td>
      <td class="CTxtContent">
        <input name="Noicap" type="text" class="CTextBoxUnder" id="Noicap" value="<%=Noicap%>" size="35" />      </td>
    </tr>
    <tr>
      <td class="CTxtContent"  align="right">Trình độ học vấn: </td>
      <td class="CFontVerdana10"><span class="CTxtContent">
        <input name="Hocvan" type="text" class="CTextBoxUnder" id="Hocvan" value="<%=Hocvan%>" size="35" />
      </span></td>
    </tr>
    <tr>
      <td class="CTxtContent"  align="right">Dân tộc: </td>
      <td class="CFontVerdana10"><span class="CTxtContent">
        <input name="DanToc" type="text" class="CTextBoxUnder" id="DanToc" value="<%=DanToc%>" size="35"/>
      </span></td>
    </tr>
    <tr>
      <td class="CTxtContent"  align="right">Điện thoại:</td>
      <td class="CTxtContent"><span class="CFontVerdana10">
        <input name="Tel" type="text" class="CTextBoxUnder" id="Tel" value="<%=Tel%>" size="35"/>
      </span></td>
    </tr>
    <tr>
      <td class="CTxtContent"  align="right">Mobile:</td>
      <td class="CFontVerdana10"><input name="Mobile" type="text" class="CTextBoxUnder" id="Mobile" value="<%=Mobile%>" size="35"/></td>
    </tr>
    
    <tr>
      <td class="CTxtContent"  align="right">Email:</td>
      <td class="CTxtContent"><input name="Email" type="text" class="CTextBoxUnder" value="<%=Email%>" size="35" /></td>
    </tr>
    <tr>
      <td height="56"  align="right" class="CTxtContent">Quê quán:</td>
      <td class="CTxtContent">
        <textarea name="Cu_tru" cols="35" rows="3" id="Cu_tru" class="CTextBoxUnder"><%=Cu_tru%></textarea>      </td>
      <td width="190"  align="right" class="CTxtContent" >Địa chỉ liên hệ: </td>
      <td width="350" class="CTxtContent"  >
        <textarea name="Diachi" cols="35" rows="3" id="Diachi" class="CTextBoxUnder"><%=Diachi%></textarea>      </td>
    </tr>
    <tr>
      <td class="CTxtContent" align="right">Tài khoản ngân hàng: </td>
      <td class="CTxtContent"><input name="TK_NganHang" type="text" class="CTextBoxUnder" id="TK_NganHang" value="<%=TK_NganHang%>" size="35"></td>
      <td class="CTxtContent" align="right">Ngân hàng: </td>
      <td class="CTxtContent">
	  <select name="BankID" id="BankID">
          <option value="0">Mời chọn</option>
          <%
		sql = "SELECT * FROM Bank Order BY BankName"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.open sql,Con,3
		If not rs.eof Then
			Do while not rs.eof
				%>
          <option value="<%=rs("BankID")%>" <%if BankID=rs("BankID") then%> selected="selected"<%end if%> ><%=rs("BankName")%></option>
          <%
				rs.movenext
			Loop
		End If
		%>
      </select></td>
    </tr>
    <tr>
      <td class="CTxtContent"  align="right">&nbsp;</td>
      <td class="CTxtContent">&nbsp;</td>
      <td  align="right" class="CTxtContent" >&nbsp;</td>
      <td class="CTxtContent"  >&nbsp;</td>
    </tr>
    <tr>
      <td colspan="4"  align="left" class="CTieuDeNhoNho" background="../../images/bMenu.gif" height="26">&nbsp;&nbsp;&nbsp;Hoàn cảnh gia đình</td>
    </tr>
    <tr>
      <td class="CTxtContent"  align="right" valign="top">Họ tên bố mẹ:<br><font class="CSubTitle">(họ tên, năm sinh, nơi ở, nghề nghiệp)</font> </td>
      <td colspan="3" class="CTxtContent">
        <textarea name="infoBoMe" cols="75" rows="7" id="infoBoMe" class="CTextBoxUnder"><%=infoBoMe%></textarea>      </td>
    </tr>
    <tr>
      <td class="CTxtContent"  align="right" valign="top" >
	    <p>Họ tên anh chị em ruột:<br> 
	      <font class="CSubTitle">(Thông tin bao gồm: quan hệ(anh,chị,em), họ tên, năm sinh, nơi ở, nghề nghiệp, trình độ) </font></p>
      <p><font class="CSubTitle">- mỗi người một dòng</font></p></td>
      <td colspan="3" class="CTxtContent" valign="top" ><textarea name="infoAnhEm" cols="75" rows="7" id="infoAnhEm" class="CTextBoxUnder"><%=infoAnhEm%></textarea></td>
    </tr>
    <tr>
      <td class="CTxtContent"  align="right" valign="top" ><p>Họ tên chồng/vợ và các con:<br>
          <font class="CSubTitle">(Thông tin bao gồm:  họ tên, năm sinh, nơi ở, nghề nghiệp, trình độ)</font></p>
        <p><font class="CSubTitle"> - mỗi người một dòng</font> </p></td>
      <td colspan="3" class="CTxtContent" valign="top" ><textarea name="infoVoChongCon" cols="75" rows="7" id="infoVoChongCon" class="CTextBoxUnder" ><%=infoVoChongCon%></textarea></td>
    </tr>
    <tr>
      <td colspan="4"  align="left" class="CTieuDeNhoNho" background="../../images/bMenu.gif" height="26">&nbsp;&nbsp;&nbsp;Quá trình hoạt động của bản thân</td>
    </tr>	
    <tr>
      <td colspan="4" valign="top" class="CTxtContent" align="center" style="border-bottom:#CCCCCC solid 1px;"><textarea name="Hoatdongbanthan" cols="100" rows="10" id="Hoatdongbanthan" class="CTextBoxUnder"><%=Hoatdongbanthan%></textarea><br>
	  <font class="CSubTitle">Thông tin mô tả bao gồm: Từ tháng năm đến tháng năm, làm gì, ở đâu, chức vụ gì? Mỗi ý 1 dòng </font>	  </td>
    </tr>
    <tr>
      <td class="CTxtContent" align="right">&nbsp;</td>
      <td colspan="3" class="CTxtContent">&nbsp;</td>
    </tr>

    <tr>
      <td colspan="4"  align="left" class="CTieuDeNhoNho" background="../../images/bMenu.gif" height="26">&nbsp;&nbsp;&nbsp;Hợp đồng lao động<font class="CSubTitle"> - Danh cho phòng nhân sự và giám đốc</font> </td>
    </tr>	
    <tr>
      <td colspan="4" class="CTxtContent" >
	<table width="100%" border="0" cellspacing="1" cellpadding="0" class="CTxtContent">
  <tr bgcolor="#FFFFCC">
    <td ALIGN="Center" style="<%=setStyleBorder(1,1,1,1)%>" height="28"><b>STT</b></td>
    <td ALIGN="Center" style="<%=setStyleBorder(0,1,1,1)%>"><b>Ngày HĐ</b></td>
    <td ALIGN="Center" style="<%=setStyleBorder(0,1,1,1)%>"><b>Hạn HĐ</b></td>
    <td ALIGN="Center" style="<%=setStyleBorder(0,1,1,1)%>"><b>Chức danh</b></td>
    <td ALIGN="Center" style="<%=setStyleBorder(0,1,1,1)%>"><b>Loại HĐ</b></td>
    <td ALIGN="Center" style="<%=setStyleBorder(0,1,1,1)%>"><b>Lương cơ bản</b></td>
    <td ALIGN="Center" style="<%=setStyleBorder(0,1,1,1)%>"><b>Hệ số</b></td>
	<td ALIGN="Center" style="<%=setStyleBorder(0,1,1,1)%>"><b>Trách nhiệm</b></td>
	<td ALIGN="Center" style="<%=setStyleBorder(0,1,1,1)%>"><b>Đã Ký</b></td>
    <td ALIGN="Center" style="<%=setStyleBorder(0,1,1,1)%>"><b>
	<a href="contract.asp?iStatus=add&NhanvienID=<%=StaffContractID%>"><img src="../../images/icons/icon_customer.gif" width="48" height="48" border="0"></a></b></td>
  </tr>	  
	  <%
		sql = "SELECT * FROM StaffContract where NhanVienID='"& StaffContractID &"'"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.open sql,Con,3
		stt = 1
		Do while not rs.eof
	  %>
	<tr>
    <td ALIGN="Center" style="<%=setStyleBorder(1,1,0,1)%>"><%=stt%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>">
	<%
	Ngay = rs("NgayHD")
	if  isdate(Ngay) = true then
		Response.Write(Day(Ngay)&"/"&month(Ngay)&"/"&year(Ngay))
	else
		Response.Write("Thiếu dữ liệu")
	end if
	%>	</td>
    <td style="<%=setStyleBorder(0,1,0,1)%>">
	<%
	Ngay = rs("EndDate")
	if  isdate(Ngay) = true then
		Response.Write(Day(Ngay)&"/"&month(Ngay)&"/"&year(Ngay))
	else
		Response.Write("Không thời hạn")
	end if
	%>	</td>
    <td style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;<%=getchucdanh(rs("ChucDanhID"))%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>"  ><%=get_ismember(rs("isMember"))%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>" align="right"><%=Dis_str_money(rs("luong"))%> đ</td>
    <td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=GetNumeric(rs("heso"),0)%></td>
	<td align="right" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(GetNumeric(rs("trachnhiem"),0))%></td>
	<td style="<%=setStyleBorder(0,1,0,1)%>" align="center">
	 	
	<%if rs("Dongy") <> 0 then%>
		<img src="../../images/icons/icon-ok.gif" align="middle" border="0">
	<%else%>
		&nbsp;
	<%end if%></td>
    <td style="<%=setStyleBorder(0,1,0,1)%>" align="center">
		<a href="contract.asp?iStatus=view&contractid=<%=rs("ID")%>"><img src="../../images/icons/Information.png" height="32" border="0"></a>
		<a href="contract.asp?iStatus=edit&contractid=<%=rs("ID")%>"><img src="../../images/bullet277.gif" height="32" border="0"></a> </td>
	</tr>
  
	  
	  <%
	  		stt = stt+ 1
			rs.movenext  
	  	loop
		%>
		</table>	  </td>
    </tr>
    <tr>
      <td colspan="2" class="CTxtContent" align="center" valign="top"></td>
      <td colspan="2"  align="center" class="CTxtContent"  valign="top"><font class="CTieuDeNhoNho" >MẪU CHỮ KÝ</font><br>
          <%if eChuKy<>"" then%>
          <img src="<%=eChuKy%>"  border="0" align="absmiddle" >
          <input type="hidden" name="imgESigna" value="<%=eChuKy%>">
          <%end if%>
          <br>
          <input name="eChuKyFile" type="file" id="eChuKyFile" size="17">
          <br>
          <%if iStatus	=	"edit" and f_permission >= 2 and eChuKy<>"" then%>
          <input type="checkbox" name="RemoveImageeChuKy" value="1">
        Xóa chữ ký
        <%end if%>
      </td>
    </tr>
    <tr>
      <td colspan="4" class="CTxtContent" height="50">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="4" align="center">
	  <input type="button" value="    Cập nhật   " name="button" onClick="checkInput()">
      <input type="reset" name="Submit2" value="   Nhập lại   ">
	  <input type="hidden" name="iStatus" value="<%=iStatus%>">
	  <input type="hidden" name="Catup" value="staff">
	  <input type="hidden" name="StaffContractID" value="<%=StaffContractID%>">	  </td>
    </tr>	
  </table>
	</form>
<script>VISUAL=4; FULLCTRL=1;</script>
<script src='../js/quickbuild.js'></script>
<script>changetoIframeEditor(document.forms[0].infoBoMe)</script>
<script>changetoIframeEditor(document.forms[0].infoAnhEm)</script>
<script>changetoIframeEditor(document.forms[0].infoVoChongCon)</script>
<script>changetoIframeEditor(document.forms[0].Hoatdongbanthan)</script>
</body>
</html>

<script language="javascript">
function checkInput()
{
	if(document.fStaff.Ho_Ten.value =='')
	{	
		alert('Xin mời nhập họ và tên cán bộ')		
		document.fStaff.Ho_Ten.focus();
		return;
	}
	document.fStaff.submit();
}
</script>