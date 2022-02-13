<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/include/lib_ajax.asp"-->
<% 
Sub Province(ProvinceName,ProvinceID,class_,style)   
	Set rs=server.CreateObject("ADODB.Recordset")
    sql="Select * From Province Order by Orderby" 
    'Response.write sql  
	rs.Open sql, con, 1
	
    response.Write "<select data-placeholder='Chọn Tỉnh/ Thành phố'  class='"&class_&"' style='"&style_&"' name="&ProvinceName&" id="&ProvinceName&">"
	response.Write"<option value='0'>Chọn Tỉnh/ Thành phố</option>"
    Do while not rs.eof
		response.Write"<option value=""" & clng(rs("ProvinceID"))  & """"
		if clng(rs("ProvinceID"))=ProvinceID then
			response.Write(" selected")
		end if
		response.Write(">")
		response.Write(Trim(rs("NameProvince")) & "</option>")
	rs.movenext
	Loop	
    response.Write "</select>"
    rs.close
    set rs=nothing
End Sub    
%>
<% 
Sub District(DistrictName,ProvinceID,DistrictID,class_,style)   
	Set rs=server.CreateObject("ADODB.Recordset")
    sql="Select DistrictID,NameDistrict From Province_district Where ProvinceID="&ProvinceID
    Response.write ("<p style='display:none'>"&sql&"</p>")
    'Response.write sql
	rs.Open sql, con, 1

	response.Write "<select data-placeholder='Chọn Quận/ Huyện'  class='"&class_&"' style='"&style_&"' name="&DistrictName&" id="&DistrictName&">"
	Response.write "<option value='0'>Chọn quận / huyện</option>"
    Do while not rs.eof
		response.Write "<option value=""" & clng(rs("DistrictID"))  & """"
		if clng(rs("DistrictID"))=DistrictID then
			response.Write(" selected")
		end if
		response.Write(">")
		response.Write(Trim(rs("NameDistrict")) & "</option>")
	rs.movenext
	Loop	
    response.Write "</select>"
    rs.close
    set rs=nothing
End Sub    
%>

<% 
Sub Wards(WardName,DistrictID,WardID,class_,style)   
	Set rs=server.CreateObject("ADODB.Recordset")
    sql="Select wardID,wardName From Province_ward Where DistrictID="&DistrictID

    'Response.write sql
	rs.Open sql, con, 1

	response.Write "<select data-placeholder='Chọn phường xã'  class='"&class_&"' style='"&style_&"' name="&WardName&" id="&WardName&">"
	Response.write "<option value='0'>Chọn phường xã</option>"
    Do while not rs.eof
		response.Write "<option value=""" & clng(rs("wardID"))  & """"
		if clng(rs("wardID"))=WardID then
			response.Write(" selected")
		end if
		response.Write(">")
		response.Write(Trim(rs("wardName")) & "</option>")
	rs.movenext
	Loop	
    response.Write "</select>"
    rs.close
    set rs=nothing
End Sub    
%>

<% 
Sub Street(streetName,wardID,streetID,class_,style)   
	Set rs=server.CreateObject("ADODB.Recordset")
    sql="Select streetID,streetName From Province_street Where wardID="&wardID

    'Response.write sql
	rs.Open sql, con, 1

	response.Write "<select data-placeholder='Chọn phường xã'  class='"&class_&"' style='"&style_&"' name="&streetName&" id="&streetName&">"
	Response.write "<option value='0'>Chọn phường xã</option>"
    Do while not rs.eof
		response.Write "<option value=""" & clng(rs("wardID"))  & """"
		if clng(rs("streetID"))=streetID then
			response.Write(" selected")
		end if
		response.Write(">")
		response.Write(Trim(rs("streetName")) & "</option>")
	rs.movenext
	Loop	
    response.Write "</select>"
    rs.close
    set rs=nothing
End Sub    
%>