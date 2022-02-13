<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/include/func_tiny.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<%	
	iStatus = Request.QueryString("iStatus")
	select case iStatus
		case "add" 
			NewsID		=	GetNumeric(Request.QueryString("NewsID"),0)
			NewsConnectID  = GetNumeric(Request.QueryString("NewsConnectID"),0)
			if NewsID<>0 and NewsConnectID <>0 then
				strPro 	=	"INSERT INTO Connection(NewsID,NewsConnectID)"
				strPro 	=	strPro 	+ " VALUES ('"& NewsID &"','"& NewsConnectID &"')"
				Set rsPro = Server.CreateObject("ADODB.Recordset")
				Response.Write(strPro)
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
			NewsID		=	GetNumeric(Request.QueryString("NewsID"),0)
			NewsConnectID  = GetNumeric(Request.QueryString("NewsConnectID"),0)
			strPro 	=	"Delete Connection where NewsID='"&NewsID&"' and NewsConnectID="&NewsConnectID
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
		
