<%@  language="VBSCRIPT" codepage="65001" %>
<!--#include virtual="/include/config.asp" -->


<%
    getType = Trim(Request.Form("type"))
    select case getType 
        case "GetListCategory"
           call GetListCategory()
        case else
    end select
    
    
%>


<%
    sub  GetListCategory()
        categoryLoai = Trim(Request.Form("val"))
    'Response.Write categoryLoai
        Call List_Category(0, "Tất cả danh mục","VN",categoryLoai)
    
    end sub
    ' Call List_Category(0, "Tất cả danh mục","VN",6)
%>

<%Sub List_Category(CatSelect,CatTitle,LanguageId,CategoryLoai)
	'CatSelect=CategoryID is choosen.
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	
	sql="SELECT *"
	sql=sql & " FROM NewsCategory"
	if LanguageId="NONE" then
		if CategoryLoai<>-1 then
			sql=sql + " CategoryLoai = '"& CategoryLoai &"'"
		end if	
	else
		sql=sql & " Where Languageid='" & LanguageId & "'"
		if CategoryLoai<>-1 then
			sql=sql + " and CategoryLoai = '"& CategoryLoai &"'"
		end if		
	end if

	sql=sql & " ORDER BY CategoryOrder"
    'Response.Write sql
	rs.Open sql, con, 1
	
	response.Write "<select name=""categoryid"" id=""categoryid"">" & vbNewline &_
                   "<option value=""0"">" & CatTitle & "</option>" & vbNewline
    Do while not rs.eof
		response.Write "<option value=""" & rs("categoryid")  & """"
		if rs("CategoryLevel")=1 then
			 response.Write " style=""COLOR: Red; background-color:#FFFFFF"""
		end if
		
		if Cint(rs("categoryid"))=Cint(CatSelect) then
			response.Write(" selected")
		end if
		response.Write(">")
		if rs("CategoryLevel")=1 then
			response.write "&#8226;"
		else
			for i=2 to rs("CategoryLevel")
				response.Write("--")
			Next
		end if
		response.Write(rs("CategoryName") & "</option>" & vbNewline)
	rs.movenext
	Loop
	rs.close
    response.Write "</select>"
End Sub%>