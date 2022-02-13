<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
f_permission = administrator(false,session("user"),"m_order_input")
if f_permission < 3 then
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
		branchID=Clng(Trim(Request.Form("BranchID"&j)))
		BranchName=Trim(Request.Form("txtBranchName"&j))
		Tel=Trim(Request.Form("txtTel"&j))
		address=Trim(Request.Form("txtAddress"&j))
		mobile=Trim(Request.Form("txtMobile"&j))
		sqlBranch 	=	"update Branch set BranchName=N'"& BranchName &"', tel =N'"& Tel &"', address = N'"& address &"', mobile = N'"& mobile &"'"
		sqlBranch = sqlBranch + " where branchID = " & branchID
		Set rsBranch = Server.CreateObject("ADODB.Recordset")
		rsBranch.open sqlBranch,con,1
		set	rsBranch = nothing
	next
	i=i+1
	BranchName=Trim(Request.Form("txtBranchName"&i))
	if  BranchName <>"" or BranchName <> NULL then
		Tel=Trim(Request.Form("txtTel"&i))
		address=Trim(Request.Form("txtAddress"&i))
		mobile=Trim(Request.Form("txtmobile"&i))
		sqlBranch =	"insert into Branch(BranchName, tel, address, mobile) values(N'"& BranchName &"',N'"& Tel &"',N'"& address &"',N'"& Mobile &"')"
		Set rsBranch = Server.CreateObject("ADODB.Recordset")
		rsBranch.open sqlBranch,con,1
		set	rsBranch = nothing
	end if
	Response.Redirect("branch_list.asp")
 case else
 	action = Request.QueryString("action")
	If action = "del" Then
		BranchID = Request.QueryString("BranchID")
		sqlCheck = "SELECT * FROM Nhanvien WHERE BranchID = " & BranchID
		Set rsCheck = Server.CreateObject("ADODB.Recordset")
		rsCheck.open sqlCheck,Con,1
		If not rsCheck.eof Then
			Response.Write("Vẫn tồn tại nhân viên thuộc chi nhánh này <br>")
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
			sql="delete Branch where BranchID=" & BranchID
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
