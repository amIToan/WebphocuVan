<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Call PhanQuyen("QLyNhanVien")%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/md5.asp" -->
<%
	Dim rs
	Set rs=Server.CreateObject("ADODB.Recordset")
	UserName=Request.QueryString("param")
	UserName=Replace(UserName,"'","''")
		
	if 	request.Form("action")="Update" then
		booAction=True
		booError=False
		UserPwd=Trim(replace(request.Form("UserPwd"),"'","''"))
		UserPwdCon=Trim(replace(request.Form("UserPwdCon"),"'","''"))
		UserEmail=Trim(replace(request.Form("UserEmail"),"'","''"))
		UserFullName=Trim(replace(request.Form("UserFullName"),"'","''"))
		UserTitle=Trim(replace(request.Form("UserTitle"),"'","''"))
		UserOnline=1
		UserOnline=request.Form("UserOnline")
		
		If UserEmail="" then
			sUserEmail="Bắt buộc"
			booError=True
		end if
		If UserPwd<>UserPwdCon  then
			sUserPwdCon="Không khớp"
			booError=True
		end if
		
		If Userpwd<>"" and len(UserPwd)<6 then
			sUserPwd="> 6 ký tự"
			booError=True
		end if
		
		if not booError then
			sql="Update [User] set "
			sql=sql & "UserEmail='" & UserEmail & "'"
			sql=sql & ",UserFullName=N'" & UserFullName & "'"
			sql=sql & ",UserTitle=N'" & UserTitle & "'"
			sql=sql & ",UserOnline='" & UserOnline & "'"			
			if UserPwd<>"" then
				sql=sql & ",UserPwd='" & md5(UserPwd) & "'"
			end if
			sql=sql & " where username=N'" & username & "'"

			rs.open sql,con,1
			set rs=nothing

			response.Write "<script language=""JavaScript"">" & vbNewline &_
				"<!--" & vbNewline &_
					"window.opener.location.reload();" & vbNewline &_
					"window.close();" & vbNewline &_
				"//-->" & vbNewline &_
				"</script>" & vbNewline
		end if
	else
		
		sql="select * from [User] where Username=N'" & UserName & "'"

		rs.open sql,con,1
			UserEmail=Trim(rs("UserEmail"))
			UserFullName=Trim(rs("UserFullName"))
			UserTitle=Trim(rs("UserTitle"))
			UserOnline=rs("UserOnline")
		rs.close
		set rs=nothing
	end if
%>
<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="fNew" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>?<%=Request.ServerVariables("QUERY_STRING")%>">
  <table width="100%" border="0" cellspacing="2" cellpadding="2">
    <tr align="center" valign="middle"> 
      <td height="40" colspan="2" valign="middle"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Sửa thông tin User</strong></font></td>
    </tr>
    <tr> 
      <td align="left"><font size="2" face="Arial, Helvetica, sans-serif">Tên 
        truy nhập:</font></td>
      <td align="left"><font size="2" face="Arial, Helvetica, sans-serif"><b><%=Username%></b></font></td>
    </tr>
    <tr> 
      <td align="left"><font size="2" face="Arial, Helvetica, sans-serif">Email:</font></td>
      <td align="left"><input name="UserEmail" type="text" id="UserEmail" size="30" maxlength="50" value="<%=UserEmail%>">
        <font size="1" face="Arial, Helvetica, sans-serif">(<font color="#FF0000"><strong>*<%=sUserEmail%></strong></font>)</font> 
      </td>
    </tr>
    <tr> 
      <td align="left"><font size="2" face="Arial, Helvetica, sans-serif">Họ tên:</font></td>
      <td align="left"><input name="UserFullName" type="text" id="UserFullName" size="30" maxlength="50" value="<%=UserFullName%>"></td>
    </tr>
    <tr> 
      <td align="left"><font size="2" face="Arial, Helvetica, sans-serif">Chức 
        danh: </font></td>
      <td align="left"><font size="2" face="Arial, Helvetica, sans-serif">
        <input name="UserTitle" type="text" id="UserTitle" size="30" maxlength="50" value="<%=UserTitle%>">
        </font></td>
    </tr>
    <tr> 
      <td align="left"><font size="2" face="Arial, Helvetica, sans-serif">Mật 
        khẩu:</font></td>
      <td align="left"><input name="UserPwd" type="password" id="UserPwd" size="30" maxlength="30">
        <font size="1" face="Arial, Helvetica, sans-serif">(<font color="#FF0000"><strong>*<%=sUserPwd%></strong></font>)</font></td>
    </tr>
    <tr> 
      <td align="left"><font size="2" face="Arial, Helvetica, sans-serif">Gõ lại 
        MK: </font></td>
      <td align="left"><input name="UserPwdCon" type="password" id="UserPwdCon" size="30" maxlength="30">
        <font size="1" face="Arial, Helvetica, sans-serif">(<font color="#FF0000"><strong>*<%=sUserPwdCon%></strong></font>)</font> 
      </td>
    </tr>
    <tr> 
      <td align="left"><font size="2" face="Arial, Helvetica, sans-serif">Online:</font></td>
      <td align="left">
	      <select size="1" name="UserOnline">
				  <option value="1" <%if UserOnline=1 then %> selected <%End if%>>On</option>
				  <option value="0" <%if UserOnline=0 then %> selected <%End if%>>Off</option>
		  </select><font size="1" face="Arial, Helvetica, sans-serif">(<font color="#FF0000"><strong>Chỉ có tác dụng trong phần giao lưu</strong></font>)</font> 
      </td>
    </tr>
    <tr> 
      <td align="center" colspan="2" height="35" valign="bottom"> <input type="submit" name="Submit" value="Sửa"> 
        <input type="button" name="Submit2" value="Đóng cửa sổ" onClick="javascript: window.close();"> 
		<input type="hidden" name="action" value="Update">
      </td>
    </tr>
  </table>
</form>
</body>
</html>
