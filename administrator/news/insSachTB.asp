<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/include/func_tiny.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<%
f_permission = administrator(false,session("user"),"m_sys")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%	
	iStatus = Request.QueryString("iStatus")
	select case iStatus
		case "add" 
			NewsID		=	GetNumeric(Request.QueryString("NewsID"),0)
			IDSachTB  = GetNumeric(Request.QueryString("IDSachTB"),0)
			if NewsID<>0 and IDSachTB <>0 then
				strPro 	=	"INSERT INTO SubSachTB(IDSachTB,NewsID)"
				strPro 	=	strPro 	+ " VALUES ('"& IDSachTB &"','"& NewsID &"')"
				Set rsPro = Server.CreateObject("ADODB.Recordset")
				rsPro.open  strPro,Con,1
				set rsPro = nothing
			end if
			%>
				<script language="JavaScript">
					window.history.back();
					window.opener.location.reload();
				</script>			
			<%			
		case "Del"
			ID	=	GetNumeric(Request.QueryString("ID"),0)
			IDSachTB  = GetNumeric(Request.QueryString("IDSachTB"),0)
			strPro 	=	"Delete SubSachTB"
			if IDSachTB = 0 and ID <>0 then
				strPro 	=	strPro 	+ " where id="&ID
			elseif IDSachTB <> 0 and ID  = 0 then
				strPro 	=	strPro 	+ " where idSachTB="&IDSachTB
			end if
			Set rsPro = Server.CreateObject("ADODB.Recordset")
			rsPro.open  strPro,Con,1
			set rsPro = nothing
			%>
				<script language="JavaScript">
					window.close();
					window.opener.location.reload();
				</script>			
			<%
		end select
%>
		
