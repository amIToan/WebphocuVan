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

	if Request.Form("action")="InsertItem" then
		ItemTitle=Trim(Request.Form("ItemTitle"))
		ItemTitle=replace(ItemTitle,"'","''")
		ItemTitle=replace(ItemTitle,"""","&quot;")
		
		sql="Insert VoteItem (VoteId,ItemTitle) values"
		sql=sql & " (" & VoteId
		sql=sql & ",N'" & ItemTitle & "')"
		rs.open sql,con,1
	elseif Request.Form("action")="RemoveItem" then
		ItemId=Request.Form("ItemId")
		if isnumeric(ItemId) then
			ItemId=CLng(ItemId)
			sql="delete VoteItem where itemId=" & ItemId
			rs.open sql,con,1
		end if
	elseif Request.Form("action")="EditItem" then
		ItemId=Request.Form("ItemId")
		if isnumeric(ItemId) then
			ItemId=CLng(ItemId)
			sql="select ItemTitle from VoteItem where itemId=" & ItemId
			rs.open sql,con,1
			ItemTitle_Edited=trim(rs("ItemTitle"))
			rs.close
		end if
	elseif Request.Form("action")="UpdateItem" then
		ItemId=Request.Form("ItemId")
		if isnumeric(ItemId) then
			ItemId=CLng(ItemId)
			ItemTitle=Trim(Request.Form("ItemTitle"))
			ItemTitle=replace(ItemTitle,"'","''")
			ItemTitle=replace(ItemTitle,"""","&quot;")
			sql="Update VoteItem set ItemTitle=N'" & ItemTitle & "' where itemId=" & ItemId
			'response.write sql
			rs.open sql,con,1
		end if
	end if

	sql="SELECT * from VoteItem where VoteId=" & VoteId
	rs.open sql,con,1
	VoteTitle=GetVoteTitle(VoteId)
%>
<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="fInsertItem" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>?<%=Request.ServerVariables("QUERY_STRING")%>">
  <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr align="center" valign="middle"> 
      <td height="35" colspan="3"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Thêm 
        bớt lựa chọn cho thăm dò ý kiến</strong></font><font size="3" face="Arial, Helvetica, sans-serif"><strong><br>
        </strong> <em><font size="2">&quot;<%=VoteTitle%>&quot; </font></em> </font></td>
  </tr>
  <%i=0
  Do while not rs.eof
  i=i+1%>
  <tr> 
      <td width="5%" align="right"><font size="2" face="Arial, Helvetica, sans-serif"><%=i%>.</font></td>
    <td><font size="2" face="Arial, Helvetica, sans-serif"><%=rs("ItemTitle")%></font></td>
      <td width="10%" align="center" valign="middle"> <a href="javascript: EditItem(<%=rs("ItemId")%>);"><img src="../images/icon_edit_topic.gif" alt="Edit Item" width="15" height="15" border="0" align="absmiddle"></a> 
        <a href="javascript: RemoveItem(<%=rs("ItemId")%>);"><img src="../images/icon-recycle.gif" alt="Remove Item" width="16" height="16" border="0" align="absmiddle"></a> 
      </td>
  </tr>
  <%rs.movenext
  Loop
  rs.close
  set rs=nothing%>
  <tr> 
    <td colspan="3"><hr size="1"></td>
  </tr>
  <tr> 
    <td colspan="3">
        <input name="ItemTitle" type="text" id="ItemTitle" size="43" maxlength="200" value="<%=ItemTitle_Edited%>">
		<%if Request.Form("action")="EditItem" then%>
			<input type="Button" name="Button" value=" Sửa " onClick="javascript: UpdateItem(<%=ItemId%>);">
		<%else%>
	        <input type="Button" name="Button" value="Tạo mới" onClick="javascript: InsertItem();">
		<%end if%>
		<input type="hidden" name="action" value="InsertItem">
		<input type="hidden" name="ItemId" value="">
	</td>
  </tr>
    <tr align="center"> 
      <td colspan="3"> <a href="javascript: window.close();"><font size="2" face="Arial, Helvetica, sans-serif">Đóng 
        cửa sổ</font></a></td>
  </tr>
</table>
</form>
<script language="JavaScript">
	function InsertItem()
	{
		document.fInsertItem.action.value="InsertItem";
		//alert (document.fInsertItem.action.value);
		document.fInsertItem.submit();
	}
	function RemoveItem(theItemValue)
	{
		document.fInsertItem.action.value="RemoveItem";
		document.fInsertItem.ItemId.value=theItemValue;
		document.fInsertItem.submit();
	}
	function EditItem(theItemValue)
	{
		document.fInsertItem.action.value="EditItem";
		document.fInsertItem.ItemId.value=theItemValue;
		document.fInsertItem.submit();
	}
	function UpdateItem(theItemValue)
	{
		document.fInsertItem.action.value="UpdateItem";
		document.fInsertItem.ItemId.value=theItemValue;
		document.fInsertItem.submit();
	}
</script>
</body>
</html>
