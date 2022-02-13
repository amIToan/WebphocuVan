<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
f_permission = administrator(false,session("user"),"m_accounting")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
	action =trim(Request.Form("action"))
	action=replace(action,"'","''")
	if action = "" or action=Null then
		action =trim(Request.Form("action"))
		action=replace(action,"'","''")
	end if
%>
<html>
<head>

<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="../../css/styles.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<p>
  <%
select case action
  case "update"
	i = Request.Form("iCount")-1
	j=1
	for j= 1 to i
		BankID=Clng(Trim(Request.Form("BankID"&j)))
		BankName=Trim(Request.Form("txtBankName"&j))
		Tel=Trim(Request.Form("txtTel"&j))
		address=Trim(Request.Form("txtAddress"&j))
		sqlBank 	=	"update Bank set BankName=N'"& BankName &"', tel =N'"& Tel &"', address = N'"& address &"'"
		sqlBank = sqlBank + " where BankID = " & BankID
		Set rsBank = Server.CreateObject("ADODB.Recordset")
		rsBank.open sqlBank,con,1
		set	rsBank = nothing
	next
	i=i+1
	BankName=Trim(Request.Form("txtBankName"&i))
	if  BankName <>"" or BankName <> NULL then
		Tel=Trim(Request.Form("txtTel"&i))
		address=Trim(Request.Form("txtAddress"&i))
		mobile=Trim(Request.Form("txtmobile"&i))
		sqlBank =	"insert into Bank(BankName, tel, address) values(N'"& BankName &"',N'"& Tel &"',N'"& address &"')"
		Set rsBank = Server.CreateObject("ADODB.Recordset")
		rsBank.open sqlBank,con,1
		set	rsBank = nothing
	end if
	Response.Redirect("Bank_list.asp")
 case else
 	action = Request.QueryString("action")
	If action = "del" Then
		BankID = Request.QueryString("BankID")
		sqlCheck = "SELECT * FROM Nhanvien WHERE BankID = " & BankID
		Set rsCheck = Server.CreateObject("ADODB.Recordset")
		rsCheck.open sqlCheck,Con,1
		If not rsCheck.eof Then
			Response.Write("Vẫn tồn tại nhân viên thuộc ngân hàng này <br>")
			Response.Write("Thao tác không thực hiện được.")
									%>
			<script language="javascript">
					<!--
					parselimit = 5;
					function begintimer()
					{
						if (parselimit==1)
						{
							window.close();
						}
						else
						{
							parselimit--;
							setTimeout("begintimer()",1000)
						}
					}
					begintimer();
					//-->
			</script>
			<%
		Else
			sql="delete Bank where BankID=" & BankID
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.open sql,con,1
			set	rs = nothing
				%>
	<script language="javascript">
		window.close();
		window.opener.location.reload();
	</script>
	<%
		End If
	End If
end select
%>
</p>
</body>
</html>