<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<%
kt	= true
f_permission = administrator(false,session("user"),"m_user")
if f_permission = 0 then
	kt	=	false 
end if
f_permission = administrator(false,session("user"),"m_human")
if f_permission = 0 then
	kt	=	false 
end if
f_permission = administrator(false,session("user"),"adm")
if f_permission = 0 then
	kt	=	false 
end if
f_permission = administrator(false,session("user"),"m_editor")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
if kt = false then
	'response.Redirect("addeditStaff.asp?StaffContractID="&get_staffid(session("user"))&"&iStatus=edit")
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>

<title>xseo - Danh sách nhân viên</title>
<link href="../../css/styles.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
  <%
	img	="../../images/icons/Meber.gif"
	Title_This_Page="Danh sách  -> Cán bộ công ty"
	Call header()
	Call Menu()

	Ho_ten	=	Trim(Request.Form("Ho_ten"))
	isMember=	GetNumeric(Request.Form("CategoryStaff"),10)
	ChucdanhID=	GetNumeric(Request.Form("ChucdanhID"),0)
	PhongID=	GetNumeric(Request.Form("PhongID"),0)
	
%>
<form name="form1" method="post" action="stafflist.asp?searching=search">	
<table width="350" border="0" align="center" class="CTxtContent" style="border:#3366FF solid 1px;">
  <tr>
    <td colspan="2" style="background-color:#FFFF99" class="CTieuDeNho">THỐNG KÊ DANH SÁCH CÁN BỘ </td>
    </tr>
  <tr>
    <td>Họ tên: </td>
    <td><input name="Ho_ten" type="text" id="Ho_ten" value="<%=Ho_ten%>" size="35"></td>
  </tr>
  <tr>
    <td width="83">Loại cán bộ: </td>
    <td width="207">
  <select name="CategoryStaff">
		<option value="1" <%if isMember=1 then%> selected="selected"<%end if%>>Đang hoạt động</option>
		<option value="0" <%if isMember=0 then%> selected="selected"<%end if%>>Đã nghỉ việc</option>
		<option value="5" <%if isMember=5 then%> selected="selected"<%end if%>>Hết hợp đồng</option>
		<option value="2" <%if isMember=2 then%> selected="selected"<%end if%>>Bán thời gian</option>
		<option value="3" <%if isMember=3 then%> selected="selected"<%end if%>>Đối tác</option>
		<option value="4" <%if isMember=4 then%> selected="selected"<%end if%>>Khác</option>
		<option value="10"<%if isMember=10  then%> selected="selected"<%end if%>>Tất cả</option>
  </select>	</td>
  </tr>
  <tr>
    <td>Chức danh: </td>
    <td><%call SelectChucDanh(ChucdanhID,"ChucdanhID")%></td>
  </tr>
  <tr>
    <td>Phòng: </td>
    <td><%call selectroom(PhongID,"PhongID")%>
	</td>
  </tr>
  <tr>
    <td colspan="2" align="center"><input type="submit" name="Submit" value="  Thống kê  ">
      <input type="button" name="Submit2" value="Thêm cán bộ"  onClick="javascript: window.location = 'addeditStaff.asp?iStatus=add'"></td>
    </tr>    
</table>
</form>


<%

if Trim(Request.QueryString("searching")) = "search" and (isMember<>10 or ChucdanhID<>0 or PhongID<>0) then

sql = "SELECT n.*, Nhanvien.Ho_Ten,Nhanvien.Mobile,Nhanvien.NgaySinh FROM StaffContract AS n INNER JOIN Nhanvien ON n.NhanVienID = Nhanvien.NhanVienID WHERE     (n.NhanVienID > 0)"
if isMember=0 then
	sql = sql + " and n.isMember='"& isMember &"' or n.EndDate<'"& now &"'"
elseif isMember=2 or isMember=3 or isMember=4 then
	sql = sql + " and n.isMember='"& isMember &"'"
elseif isMember = 1 then
	sql = sql + " and n.isMember='"& isMember &"' and n.EndDate>'"& now &"' "
elseif isMember = 5 then
	sql = sql + " and n.EndDate<'"& now &"'"
end if
if Ho_ten<>"" then
	sql = sql + " and {fn UCASE(Nhanvien.Ho_Ten)} like N'%"& UCase(Ho_Ten) &"%'"
end if
if ChucdanhID > 0 then
	sql = sql + " and n.ChucdanhID='"& ChucdanhID &"'"
end if
if PhongID > 0 then
	sql = sql + " and n.PhongID='"& PhongID &"'"
end if
sql	=	sql + "  ORDER BY PhongID"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql,Con,3
If not rs.eof Then
%>
<table width="95%" align="center" cellpadding="0" cellspacing="1">
<tr>
	<td width="59" align="center" bgcolor="#FFFFCC" class="CTxtContent"  style="<%=setStyleBorder(1,1,1,1)%>"><strong>STT</strong></td>
	<td width="262" align="center" bgcolor="#FFFFCC" class="CTxtContent" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Họ và tên</strong></td>
	<td width="148" align="center" bgcolor="#FFFFCC" class="CTxtContent" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Loại cán bộ </strong></td>
	<td width="163" align="center" bgcolor="#FFFFCC" class="CTxtContent" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Chức danh</strong></td>
	<td width="147" align="center" bgcolor="#FFFFCC" class="CTxtContent" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Di động</strong></td>
	<td width="109" align="center" bgcolor="#FFFFCC" class="CTxtContent" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Ngày sinh </strong></td>
	<td width="101" align="center" bgcolor="#FFFFCC" class="CTxtContent" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Ngày hợp đồng </strong></td>
	<td width="99" align="center" bgcolor="#FFFFCC" class="CTxtContent" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Hết hợp đồng </strong></td>
	<td width="69" align="center" bgcolor="#FFFFCC" class="CTxtContent" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Cập nhật</strong></td>
</tr>
<%
	i = 1
	Do while not rs.eof
		%>
		<tr <%if i mod 2=0 then response.Write("bgcolor=""#E6E8E9""") else response.Write("bgcolor=""#FFFFFF""")%>>
			<td style="<%=setStyleBorder(1,1,0,1)%>"><%=i%></td>
			<td style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("HO_ten")%></td>
			<td style="<%=setStyleBorder(0,1,0,1)%>">
			<%
			if rs("EndDate") < Now() then
				Response.Write("Hết hạn hợp đồng")
			else
				Response.Write(get_ismember(rs("isMember")))
			end if
			%>			</td>
			<td style="<%=setStyleBorder(0,1,0,1)%>">
			<%
			sql = "SELECT * FROM Chucdanh  where ChucdanhID='"& rs("ChucdanhID") &"'"
			Set rsTemp = Server.CreateObject("ADODB.Recordset")
			rsTemp.open sql,Con,3
			if not rsTemp.eof then
				Response.Write(rsTemp("Description"))
			else
				Response.Write("&nbsp;")	
			end if
			%>

			</td>
			<td style="<%=setStyleBorder(0,1,0,1)%>">
				<%
					if rs("Mobile") <> "" or rs("Mobile") <> null then 
						
						Response.Write(rs("Mobile"))
					else
						Response.Write("&nbsp;")
					end if
				%>			</td>
			<td style="<%=setStyleBorder(0,1,0,1)%>"><%=Day(rs("NgaySinh"))%>/<%=Month(rs("NgaySinh"))%>/<%=Year(rs("NgaySinh"))%></td>
			<td style="<%=setStyleBorder(0,1,0,1)%>" ><%=Day(rs("NgayHD"))%>/<%=Month(rs("NgayHD"))%>/<%=Year(rs("NgayHD"))%></td>
			<td style="<%=setStyleBorder(0,1,0,1)%>" ><%=Day(rs("EndDate"))%>/<%=Month(rs("EndDate"))%>/<%=Year(rs("EndDate"))%></td>
			<td style="<%=setStyleBorder(0,1,0,1)%>" align="center">
			<a href="addeditStaff.asp?StaffContractID=<%=rs("NhanVienID")%>&iStatus=edit" target="_parent">
			<img src="../images/icon_edit_topic.gif" width="15" height="15" border="0" align="absmiddle"></a>
			<%if f_permission > 3 then %>
			<img src="../../images/icons/icon_closed_topic.gif" width="15" height="15"  border="0" align="absmiddle"  onClick="javascript: yn = confirm('Bạn có chắc chắn muốn xóa nhân viên này không?'); if(yn) {window.location = 'staffdelete.asp?nhanvienID=<%=rs("NhanvienID")%>'}" style="cursor:pointer;">
			<%end if%>			</td>
		</tr>

		<%
		i = i + 1
		rs.movenext
	Loop
	%>
			<tr >
		  <td>&nbsp;</td>
		  <td colspan="7" class="CTextStrong">&nbsp;</td>
		  <td>Thêm</td>
	    </tr>
</table>
	<%
	Else
		Response.Write("Không có kế quả.")
End If
elseif Trim(Request.QueryString("searching")) = "search" and (isMember=10 and ChucdanhID=0 and PhongID=0) then
sql = "SELECT * FROM Nhanvien  WHERE NhanVienID > 0"
if Ho_ten<>"" then
	sql = sql + " and {fn UCASE(Nhanvien.Ho_Ten)} like N'%"& UCase(Ho_Ten) &"%'"
end if
sql	=	sql + " ORDER BY Ho_ten"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql,Con,3
If not rs.eof Then
%>
<table width="95%" align="center" cellpadding="0" cellspacing="1">
<tr>
	<td width="59" align="center" bgcolor="#FFFFCC" class="CTxtContent"  style="<%=setStyleBorder(1,1,1,1)%>"><strong>STT</strong></td>
	<td width="262" align="center" bgcolor="#FFFFCC" class="CTxtContent" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Họ và tên</strong></td>
	<td width="163" align="center" bgcolor="#FFFFCC" class="CTxtContent" style="<%=setStyleBorder(0,1,1,1)%>"><strong>CMND</strong></td>
	<td width="147" align="center" bgcolor="#FFFFCC" class="CTxtContent" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Di động</strong></td>
	<td width="109" align="center" bgcolor="#FFFFCC" class="CTxtContent" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Ngày sinh </strong></td>
	<td width="69" align="center" bgcolor="#FFFFCC" class="CTxtContent" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Cập nhật</strong></td>
</tr>
<%
	i = 1
	Do while not rs.eof
		%>
		<tr <%if i mod 2=0 then response.Write("bgcolor=""#E6E8E9""") else response.Write("bgcolor=""#FFFFFF""")%>>
			<td style="<%=setStyleBorder(1,1,0,1)%>"><%=i%></td>
			<td style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("Ho_ten")%><br>
			Đ/c: <%=rs("Diachi")%>
			</td>
			<td style="<%=setStyleBorder(0,1,0,1)%>"><%=rs("CMT")%></td>
			<td style="<%=setStyleBorder(0,1,0,1)%>">
				<%
					if rs("Mobile") <> "" or rs("Mobile") <> null then 
						
						Response.Write(rs("Mobile"))
					else
						Response.Write("&nbsp;")
					end if
				%>			</td>
			<td style="<%=setStyleBorder(0,1,0,1)%>"><%=Day(rs("NgaySinh"))%>/<%=Month(rs("NgaySinh"))%>/<%=Year(rs("NgaySinh"))%></td>
			<td style="<%=setStyleBorder(0,1,0,1)%>" align="center">
			<a href="addeditStaff.asp?StaffContractID=<%=rs("NhanVienID")%>&iStatus=edit" target="_parent">
			<img src="../images/icon_edit_topic.gif" width="15" height="15" border="0" align="absmiddle"></a>
			<%if f_permission > 3 then %>
			<img src="../../images/icons/icon_closed_topic.gif" width="15" height="15"  border="0" align="absmiddle"  onClick="javascript: yn = confirm('Bạn có chắc chắn muốn xóa nhân viên này không?'); if(yn) {window.location = 'staffdelete.asp?nhanvienID=<%=rs("NhanvienID")%>'}" style="cursor:pointer;">
			<%end if%>			</td>
		</tr>

		<%
		i = i + 1
		rs.movenext
	Loop
	%>
			<tr >
		  <td>&nbsp;</td>
		  <td colspan="4" class="CTextStrong">&nbsp;</td>
		  <td>Thêm</td>
	    </tr>
</table>
	<%
	Else
		Response.Write("Không có kế quả.")
End If
end if
%>


</body>
</html>
