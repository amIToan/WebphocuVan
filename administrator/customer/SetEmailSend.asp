<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<%		
	all			=	Trim(Request.QueryString("All"))
	if all <> "ok" then
		iSoEmail	=	GetNumeric(Request.Form("iSoEmail"),0)
	end if
	ST= 1
%>
<%
Set arEmail = Server.CreateObject("Scripting.Dictionary")
Set arName = Server.CreateObject("Scripting.Dictionary")
if all <> "ok" then
	for i = 0 to iSoEmail 
		if GetNumeric(Request.Form("CbEmailKhach"&i),0) = 1 then
			IDEmail		=	GetNumeric(Request.Form("IDEmail"&i),0)
			m_sql	=	"Select Ten,Email from Email where ID="&IDEmail&" and Disabled<>0"
			set m_rs	=	Server.CreateObject("ADODB.recordset")
			m_rs.open m_sql,con,1
			if not m_rs.eof then
				Email		=	m_rs("Email")
				Ten			=	m_rs("Ten")
			end if
			set m_rs	=	nothing
			if Xungho = "" then
				'On Error Resume Next 
				Xungho = fXungHo(Ten)
			end if
			if (arEmail.Exists(IDEmail) = false) and (IDEmail<>0) Then				
				arEmail.Add IDEmail, Email
				arName.Add IDEmail, Ten
				ST	=	ST + 1
			end if			
		end if
	next	
else
	sql = "Select * From Email where Disabled<>0"
	set rsEmail	=	Server.CreateObject("ADODB.recordset")
	rsEmail.open sql,con,1
	do while not rsEmail.eof 
		IDEmail		=	rsEmail("ID")
		Email		=	rsEmail("Email")
		Ten			=	rsEmail("Ten")
			
		If (arEmail.Exists(IDEmail) = false) and (IDEmail<>0) Then				
			arEmail.Add IDEmail, Email
			arName.Add IDEmail, Ten							
			ST	=	ST + 1	
		end if							
		rsEmail.movenext
	loop
					
end if
Set Session("arEmail") = arEmail
Set Session("arName") = arName
set arEmail	=	nothing
set arName	=	nothing
Response.Redirect("SendMailTT.asp")
%>			