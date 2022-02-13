<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<link href="../../css/styles.css" rel="stylesheet" type="text/css" />
<%
f_permission = administrator(false,session("user"),"m_customer")
if f_permission < 3 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
	txtPassword		=	Trim(Request.Form("txtPassword"))
    txtname		    =	Trim(Request.Form("txtname"))
    txtserverIP		=	Trim(Request.Form("txtserverIP"))
	txtHoTen		=	Trim(Request.Form("txtHoTen"))
	txtForm     	= 	Trim(Request.Form("txtForm"))
	txtTieuDe  		= 	Uni2NONE(Trim(Request.Form("txtTieuDe")))
	txtNoiDung  	= 	Trim(Request.Form("txtNoiDung"))
	set arEmail		=	Session("arEmail")
	set arName		=	Session("arName")	 

    if txtForm<>"" and txtTieuDe<>"" and txtNoiDung<>"" then
		stt = 0
		For Each Key in arEmail	
    '		Set Mail 		= Server.CreateObject("Persits.MailSender")
    '		Mail.Host 		= txtserverIP
    '		Mail.From 		= txtForm  ' From address
    '		Mail.FromName 	= txtname
		    Ten			=	arName(Key)
		    TieuDe=""
		    NoiDung	=	""
            if Ten <> "" and txtHoTen <>"" then
		        TieuDe	=	Uni2NONE(replace(txtTieuDe,txtHoTen,Ten))
            end if
    '		Mail.Subject = TieuDe
            if Ten <> "" and txtHoTen <>"" then
		        NoiDung	=	replace(txtNoiDung,txtHoTen,Ten)
            end if
            NoiDung =   Replace(NoiDung,"/images_upload/",Request.ServerVariables("SERVER_NAME")&"/images_upload/")
    '		Mail.Body = NoiDung
    '		Mail.IsHTML = True 
    '		Mail.ContentTransferEncoding = "Quoted-Printable"
    '		Mail.CharSet = "UTF-8"
    '						 
    '		Mail.AddAddress arEmail(Key)
    '		On Error Resume Next ' catch errors
    '		Mail.Send	' send message		
            mailBody  =  NoiDung
            '19 - 06 - 2015
            Dim ObjSendMail
            Set ObjSendMail =server.createobject("CDO.Message")
   
            'This section provides the configuration information for the remote SMTP server.
            ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2'Send the message using the network (SMTP over the network).
            ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = txtserverIP
            ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
            ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = 1
            ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
   
            ' Google apps mail servers require outgoing authentication. Use a valid email address and password registered with Google Apps.
            ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic (clear-text) authentication
            ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = txtForm 'your Google apps mailbox address
            ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = txtPassword 'Google apps password for that mailbox
   
            ObjSendMail.Configuration.Fields.Update
            ObjSendMail.To = arEmail(Key)

            ObjSendMail.Subject =TieuDe
            ObjSendMail.From = txtForm
            'ObjSendMail.IsHTML = True 
            ObjSendMail.BodyPart.Charset = "utf-8" 
            ObjSendMail.HTMLBody = mailBody
  
            ObjSendMail.Send
            Set ObjSendMail = Nothing 

            sql = 	"Insert into EmailCompose(FromEmail,ToEmail,TitleEmail) values('"& txtForm  &"','"& arEmail(Key)  &"',N'"& TieuDe &"')"
            set rs=server.CreateObject("ADODB.Recordset")		
            rs.open sql,con,3
            set rs=nothing	
                
	        stt = stt+1
        next 
	 else
			Response.Write("Content is not empty !")
			Response.Write("<div align=""center""><a href=""javascript:history.back()"">Back</a></div>&nbsp;&nbsp;") 	
	 end if
%>

<html>
	<head>
	<title><%=PAGE_TITLE%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
	<link href="../../css/styles.css" rel="stylesheet" type="text/css">
	</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
	img	="../../images/icons/icon_email1.gif"
	Call header()

%>

<table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td Class="CtieuDe" align="center">COMPOSE EMAIL</td>
  </tr>
  <tr>
    <td >
        <%
		Response.Write("<div align=""center"">All is "& stt &" emails send finished! <br><a href=""javascript:history.back;"">Back</a></div>&nbsp;&nbsp;") 	
        %>
	  </td>
  </tr>
  <tr>
    <td ></td>
  </tr>
</table>
</body>
</html>