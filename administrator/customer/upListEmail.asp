<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<%	


	GhiChu		=	Trim(Request.Form("txtGhichu"))
    IDCongViec		=	GetNumeric(Request.Form("sel_cong_viec"),0)

    Dim ArrEmail
    ArrEmail = Split(GhiChu,chr(13))
    stt = 0
    for i = 0 to UBound(ArrEmail)
        Email = trim(ArrEmail(i))
        itest = 1
        if Email = "" or Len(Email) < 5  then
            itest = 0
        end if

        if itest = 1 then
            sql = "select ID from Email where Email ='"& Email &"'"
            set rs1=server.CreateObject("ADODB.Recordset")		
            rs1.open sql,con,3
            if rs1.EOF then
                    stt = stt + 1
	                sql = 	"Insert into Email(IDCongViec,email) values('"& IDCongViec  &"','"& Email  &"')"
	                set rs=server.CreateObject("ADODB.Recordset")		
	                rs.open sql,con,3
	                set rs=nothing				
            end if
        end if
    next

    Response.Write("Update finished "& stt &" email")

%>		
