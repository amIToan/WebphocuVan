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
	if Request.QueryString("param")<>"" and IsNumeric(Request.QueryString("param")) then
		VoteId=Clng(Request.QueryString("param"))
	else
		response.Redirect("/administrator/default.asp")
	end if
	if Request.QueryString("CatId")<>"" and IsNumeric(Request.QueryString("CatId")) then
		CatId=Clng(Request.QueryString("CatId"))
	else
		response.Redirect("/administrator/default.asp")
	end if
	Call AuthenticateWithRole(CatId,Session("LstRole"),"ap")
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	
	'sql="Select VoteTitle from Vote where VoteId=" & VoteId
	'rs.open sql,con,1
	'	VoteTitle=Trim(rs("VoteTitle"))
	'rs.close

	sql="SELECT v.VoteStatus, v.VoteTitle, v.VoteNote, VoteTotal as Summary"
	sql=sql & "	FROM Vote v "
	sql=sql & " WHERE (v.VoteId = " & VoteId & ")"
	
	rs.open sql,con,1
		VoteTitle=Trim(rs("VoteTitle"))
		VoteStatus=Trim(rs("VoteStatus"))
		Summary=Clng(rs("Summary"))
		VoteNote=Trim(rs("VoteNote"))
	rs.close
	sql="select * from VoteItem where Voteid=" & VoteId
	rs.open sql,con,1

%>
<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="">
  <table width="98%" border="0" cellspacing="1" cellpadding="1" align="center">
    <tr> 
      <td><p align="center"><font face="Arial, Helvetica, sans-serif"><strong><font size="2"><%=VoteTitle%></font></strong></font>
	  	<%if VoteNote<>"" then%>
	  		<br><em><font size="2" face="Arial, Helvetica, sans-serif">(<%=VoteNote%>)</font></em>
		<%end if%>
	  </p></td>
    </tr>
    <tr align="left"> 
      <td><font size="2" face="Arial, Helvetica, sans-serif">Tổng cộng có: <strong><%=Summary%></strong> 
        lượt bầu chọn</font></td>
    </tr>
    <tr> 
      <td align="right"><table width="100%%" border="0" cellpadding="0" cellspacing="1" bordercolor="#000000" bgcolor="#000000">
          <tr bgcolor="#FFFFFF"> 
            <td align="center"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Tiêu 
              đề</strong></font></td>
            <td align="center"><font size="2" face="Arial, Helvetica, sans-serif">Số 
              lượt</font></td>
            <td align="center"><font size="2" face="Arial, Helvetica, sans-serif">Tỷ lệ</font></td>
          </tr>
		  <%Do while not rs.eof%>
          <tr bgcolor="#FFFFFF"> 
            <td><input type="<%if VoteStatus then%>checkbox<%else%>radio<%end if%>">
              <font size="2" face="Arial, Helvetica, sans-serif"><%=rs("ItemTitle")%></font></td>
            <td align="center"><font size="2" face="Arial, Helvetica, sans-serif"><%=rs("ItemCount")%></font></td>
            <td align="center"><font size="2" face="Arial, Helvetica, sans-serif"><%Response.Write Int((Clng(rs("ItemCount"))*100)/Summary)%>%</font></td>
          </tr>
		  <%rs.movenext
		  Loop
		  rs.close
		  set rs=nothing%>
        </table></td>
    </tr>
	 <tr> 
      <td align="center"><font size="2" face="Arial, Helvetica, sans-serif"><a href="javascript: window.close();">Đóng 
        cửa sổ</a></font></td>
    </tr>
  </table>
</form>
</body>
</html>
