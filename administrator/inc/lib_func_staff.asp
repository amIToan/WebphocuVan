<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
function getMobileStaff(ID,is_styte)
if is_styte = 0 then
	sql = "Select Tel From NhanVien n where n.NhanVienID='"&ID&"' and n.isMember='"& isMember &"' or n.EndDate<'"& now &"'"
else
	sql = "Select Mobile as Tel From NhanVien where NhanVienID='"&ID&"'"
end if
set rstemp = Server.CreateObject("ADODB.recordset")
rstemp.open sql,con,1
if not rstemp.eof then
	Tel = rstemp("Tel")
end if
set rstemp = nothing
getMobileStaff	= Tel
end function
%>