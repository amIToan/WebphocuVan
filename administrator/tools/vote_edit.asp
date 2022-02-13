<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
f_permission = administrator(false,session("user"),"m_ads")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
	Dim rs
	if Request.QueryString("param")<>"" and IsNumeric(Request.QueryString("param")) then
		VoteId=Clng(Request.QueryString("param"))
	else
		response.Redirect("/administrator/default.asp")
		response.End()
	end if
	
	if Request.Form("action")="Update" then
		IsHomeVote=Request.Form("IsHomeVote")
		if IsNumeric(IsHomeVote) and IsHomeVote<>"" then
			IsHomeVote=Clng(IsHomeVote)
		else
			IsHomeVote=0
		end if
		IsCatHomeVote=Request.Form("IsCatHomeVote")
		if IsNumeric(IsCatHomeVote) and IsCatHomeVote<>"" then
			IsCatHomeVote=Clng(IsCatHomeVote)
		else
			IsCatHomeVote=0
		end if
		VoteTitle=Trim(Replace(Request.Form("VoteTitle"),"'","''"))
		VoteTitle=Trim(Replace(VoteTitle,"""","&quot;"))
		if VoteTitle="" then
			sVoteTitle="Bắt buộc"
		end if
		VoteStatus=Request.Form("VoteStatus")
		if IsNumeric(VoteStatus) and VoteStatus<>"" then
			VoteStatus=Clng(VoteStatus)
		else
			sVoteStatus="Bắt buộc"
		end if
		CatId=Clng(Request.Form("CatId_DependRole"))
		
		'Kiểm tra xem Chuyên mục mới sửa User có quyền ap hoặc ad không?
		Call AuthenticateWithRole(CatId,Session("LstRole"),"ap")
		if CatId=0 then
			sCatId="Bắt buộc"
		end if
		LanguageId=request.Form("languageid")
		VoteNote=Trim(Replace(Request.Form("VoteNote"),"'","''"))
		VoteNote=Trim(Replace(VoteNote,"""","&quot;"))
		
		if sCatId="" and sVoteTitle="" and sVoteStatus=""then
			StatusID=GetRoleOfCat_FromListRole(CatId,Session("LstRole"))
			sql="UPDATE Vote set"
			sql=sql & " VoteTitle=N'" & VoteTitle & "'"
			sql=sql & ",CategoryId=" & CatId
			sql=sql & ",IsHomeVote=" & IsHomeVote
			sql=sql & ",IsCatHomeVote=" & IsCatHomeVote
			sql=sql & ",statusId='" & StatusID & "'"
			sql=sql & ",Languageid='" & LanguageId & "'"
			sql=sql & ",Approver=N'" & session("user") & "'"
			sql=sql & ",ApproverDate='" & now() & "'"
			sql=sql & ",VoteStatus=" & VoteStatus
			sql=sql & ",VoteNote=N'" & VoteNote & "'"
			sql=sql & " WHERE VoteId=" & VoteId
			
			set rs=server.CreateObject("ADODB.Recordset")
			rs.open sql,con,1
			'rs.open sql,con,1 : Quyen Ghi
			'rs.open sql,con,3 : Quyen Doc
			set rs=nothing
			response.Write "<script language=""JavaScript"">" & vbNewline &_
			"<!--" & vbNewline &_
				"window.opener.location.reload();" & vbNewline &_
				"window.close();" & vbNewline &_
			"//-->" & vbNewline &_
			"</script>" & vbNewline
			response.End()
		end if
	else
		sql="SELECT * from Vote where Voteid=" & Voteid
		set rs=server.createObject("ADODB.Recordset")
		rs.open sql,con,3
			CatId=Clng(rs("CategoryId"))
			VoteTitle=Trim(rs("VoteTitle"))
			if rs("IsHomeVote") then
				IsHomeVote=1
			else
				IsHomeVote=0
			end if
			if rs("IsCatHomeVote") then
				IsCatHomeVote=1
			else
				IsCatHomeVote=0
			end if
			if rs("VoteStatus") then
				VoteStatus=1
			else
				VoteStatus=0
			end if
			languageid=rs("languageid")
			VoteNote=Trim(rs("VoteNote"))
		rs.close
		Call AuthenticateWithRole(CatId,Session("LstRole"),"ap")
	end if
%>
<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="fInsertVote" method="post" action="<%=request.ServerVariables("SCRIPT_NAME")%>?<%=request.ServerVariables("QUERY_STRING")%>">
  <table width="100%" border="0" cellspacing="2" cellpadding="1">
    <tr align="center" valign="middle"> 
      <td height="30" colspan="2"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Sửa đổi thăm dò ý kiến</strong></font></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><input type="checkbox" name="IsHomeVote" value="1"<%if IsHomeVote=1 then%> checked<%end if%>>
        <font size="2" face="Arial, Helvetica, sans-serif"><strong>Trang chủ</strong></font></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><input type="checkbox" name="IsCatHomeVote" value="1"<%if IsCatHomeVote=1 then%> checked<%end if%>>
        <font size="2" face="Arial, Helvetica, sans-serif"><strong>Trang chuyên 
        mục</strong></font></td>
    </tr>
    <tr> 
      <td><font size="2" face="Arial, Helvetica, sans-serif">Tiêu đề:</font></td>
      <td><input name="VoteTitle" type="text" id="VoteTitle" size="35" maxlength="200" value="<%=VoteTitle%>"><font size="1" face="Arial, Helvetica, sans-serif">(<font color="#FF0000"><strong>*</strong><%=sVoteTitle%></font>)</font></td>
    </tr>
    <tr> 
      <td><font size="2" face="Arial, Helvetica, sans-serif">Chuyên mục:</font></td>
      <td>
      	<%Call List_Category_Depend_Role(CatId, "L&#7921;a ch&#7885;n","NONE",Session("LstRole"),"ap",0)%>
      	<font size="1" face="Arial, Helvetica, sans-serif">(<font color="#FF0000"><strong>*</strong><%=sCatId%></font>)</font>
      </td>
    </tr>
    <tr> 
      <td><font size="2" face="Arial, Helvetica, sans-serif">Ngôn ngữ:</font></td>
      <td><%Call List_Language(languageid)%></td>
    </tr>
	<tr> 
      <td><font size="2" face="Arial, Helvetica, sans-serif">Trạng thái:</font></td>
      <td height="30" valign="middle"><input name="VoteStatus" type="radio" value="0"<%if VoteStatus=0 then%> checked<%end if%>>
        <font size="2" face="Arial, Helvetica, sans-serif">Chọn một</font> 
        <input type="radio" name="VoteStatus" value="1"<%if VoteStatus=1 then%> checked<%end if%>>
        <font size="2" face="Arial, Helvetica, sans-serif">Chọn nhiều</font>
		<font size="1" face="Arial, Helvetica, sans-serif">(<font color="#FF0000"><strong>*</strong><%=sVoteStatus%></font>)</font>
	  </td>
    </tr>
    <tr> 
      <td><font size="2" face="Arial, Helvetica, sans-serif">Ghi chú:</font></td>
      <td height="30" valign="middle"><input name="VoteNote" type="text" id="VoteNote" size="35" maxlength="500" value="<%=VoteNote%>"></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td height="30" valign="middle"> 
        <input type="submit" name="Submit" value="Sửa đổi"> <input type="button" name="Submit2" value="Đóng cửa sổ" onClick="javascript: window.close();">
        <input type="Hidden" name="action" value="Update">
      </td>
    </tr>
  </table>
</form>
</body>
</html>