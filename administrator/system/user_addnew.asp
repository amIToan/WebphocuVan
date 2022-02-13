<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Call PhanQuyen("QLyNhanVien")%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/md5.asp" -->
<%
	
	if 	request.Form("action")="Insert" then
		booAction=True
		booError=False
		UserName=Trim(replace(request.Form("UserName"),"'","''"))
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
		If UserPwd<>UserPwdCon then
			sUserPwdCon="Không khớp"
			booError=True
		end if
		If len(UserPwd)<6 then
			sUserPwd="> 6 ký tự"
			booError=True
		end if
		If UserName="" then
			sUserName="Bắt buộc"
			booError=True
		elseif CheckUserExist(Username)<>0 then
			sUserName="Đang sử dụng"
			booError=True
		end if
		
		if not booError then
			sql="insert into [User] (UserName,UserPwd,UserEmail,UserFullName,UserTitle, UserOnline) values "
			sql=sql & "(N'" & UserName & "'"
			sql=sql & ",'" & md5(UserPwd) & "'"
			sql=sql & ",'" & UserEmail & "'"
			sql=sql & ",N'" & UserFullName & "'"
			sql=sql & ",N'" & UserTitle & "'"
			sql=sql & ",'" & UserOnline & "')"
			
			Dim rs
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.open sql,con,1
			
			sql="insert into UserDistribution (UserName,CategoryId,User_role) values "
			sql=sql & "(N'" & UserName & "'"
			sql=sql & "," & 0
			sql=sql & ",'ed')"

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
		UserName=""
		UserEmail=""
		UserFullName=""
		UserTitle=""
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
      <td height="40" colspan="2" valign="middle"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Tạo User mới</strong></font></td>
    </tr>
    <tr> 
      <td align="left"><font size="2" face="Arial, Helvetica, sans-serif">Tên 
        truy nhập:</font></td>
      <td align="left"> <input name="UserName" type="text" id="UserName" size="30" maxlength="30" value="<%=UserName%>">
        <font size="1" face="Arial, Helvetica, sans-serif">(<font color="#FF0000"><strong>*<%=sUserName%></strong></font>)</font></td>
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
      <td align="center" colspan="2" height="35" valign="bottom"> <input type="submit" name="Submit" value="Tạo mới"> 
        <input type="button" name="Submit2" value="Đóng cửa sổ" onClick="javascript: window.close();"> 
		<input type="hidden" name="action" value="Insert">
      </td>
    </tr>
  </table>
</form>
</body>
</html>
