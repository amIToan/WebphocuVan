<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<%
	if request.QueryString("CatId")<>"" and IsNumeric(request.QueryString("CatId")) then
		Catid=Clng(request.QueryString("CatId"))
	end if
	Call AuthenticateWithRole(CatId,Session("LstRole"),"ap")
	if request.QueryString("param")<>"" and IsNumeric(request.QueryString("param")) then
		EventId=Clng(request.QueryString("param"))
	else
		response.end
	end if
	
	Dim rs
	set rs=server.CreateObject("ADODB.Recordset")
	
	'Xử lý khi add thêm tin vào sự kiện
	if request.form("Action")="Add" then
		NewsNumber=CLng(Request.Form("NewsId").Count)
		if NewsNumber>0 then
			sql="Update News set EventId=" & EventId
			for i = 1 to NewsNumber
				if isnumeric(Request.Form("NewsId")(i)) then
				  if i=1 then
					sql=sql & " WHERE NewsId=" & Clng(Request.Form("NewsId")(i))
				  else
				  	sql=sql & " OR NewsId=" & Clng(Request.Form("NewsId")(i))
				  end if
				end if
			next
			rs.open sql,con,1
		end if
	'Xử lý khi add Remove tin khỏi sự kiện
	elseif request.form("Action")="Remove" then
		NewsId=GetNumeric(Request.form("R_NewsId"),0)
		if NewsId=0 then
			response.end
		else
			sql="Update News set EventId=0 where NewsId=" & NewsId
			rs.open sql,con,1
		end if
	end if
	
	sql="SELECT	EventName FROM	Event WHERE	(EventID = " & EventId & ")"
	rs.open sql,con,3
	
	if rs.eof then
		rs.close
		set rs=nothing
		response.end
	else
		EventName=Trim(rs("EventName"))
		rs.close
	end if
	
	keyword=ReplaceHTMLToText(Request.Form("keyword"))
%>
<html>
<head>
	<title><%=PAGE_TITLE%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="fEvent" method="post" action="<%=Request.serverVariables("Script_name")%>?<%=Request.ServerVariables("Query_string")%>">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr align="center" valign="middle"> 
      <td height="40" colspan="3"><font size="2" face="Arial, Helvetica, sans-serif"><strong><%=EventName%></strong></font></td>
    </tr>
    <tr> 
      <td width="48%"><font size="2" face="Arial, Helvetica, sans-serif">Tìm kiếm:&nbsp;</font>
      	<input name="keyword" type="text" id="keyword" size="30" value="<%=keyword%>"> 
        <input type="image" name="ButtonSearch" id="ButtonSearch" align="absmiddle" src="/administrator/images/search_bt.gif" width="23" height="23" border="0">
      </td>
      <td width="4%">&nbsp;</td>
      <td width="48%">&nbsp;</td>
    </tr>
    <tr> 
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td align="center" valign="middle" bgcolor="#E6E8E9"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Danh 
              sách tin tìm được</strong></font></td>
          </tr>
        </table></td>
      <td>&nbsp;</td>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td align="center" valign="middle" bgcolor="#E6E8E9"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Danh 
              sách tin thuộc sự kiện</strong></font></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td valign="top">
      <%if keyword<>"" then
      	sql="SELECT NewsId,Title FROM News " &_
      		"WHERE (title like '%" & keyword & "%' or subtitle like '%" & keyword & "%' or description like '%" & keyword & "%') " &_
      		"		And (eventId<>" & Eventid & ") " &_
      		"ORDER BY Newsid desc"
      	Dim rs_Search
      	set rs_Search=server.CreateObject("ADODB.Recordset")
      	
      	NEWS_PER_PAGE=10
		PAGES_PER_BOOK=7
		rs_Search.PageSize = NEWS_PER_PAGE
		rs_Search.open sql,con,3
		
if not rs_Search.eof then
	if request.Form("page")="" or not isnumeric(request.Form("page")) then
		page=1
	else
		page=Cint(request.Form("page"))
	end if
	rs_Search.AbsolutePage = CLng(page)
	i=0
%>
      	<table width="100%" border="0" cellspacing="0" cellpadding="2">
		 <%Do while not rs_Search.eof and i<rs_Search.pagesize%>
          <tr> 
            <td width="1%" valign="top"><input name="NewsId" type="checkbox" id="NewsId" value="<%=rs_Search("NewsId")%>"></td>
            <td width="99%"><div align="justify"><font size="2" face="Arial, Helvetica, sans-serif"><%=rs_Search("Title")%></font></div></td>
          </tr>
         <%i=i+1
         rs_Search.movenext
         Loop
         %>
         <tr><td colspan="2" align="center" valign="middle">
	         <p align="center"><span style="font-family: Arial; font-size: 10pt">
				<%for i=1 to rs_Search.pagecount 
					if i<>page then
						response.Write"<a href=""javascript: fSearchOnSubmit('" & i & "');"">" & i & "</a>&nbsp;|&nbsp;"
					else
						response.Write"<font color=""red""><b>" & i & "</b></font>&nbsp;|&nbsp;"
					end if
				Next%>
			</span>
         </td></tr>
        </table>
<%	rs_Search.close
    set rs_Search=nothing
end if 'if not rs_Search.eof then
		end if 'if keyword<>"" then
      %>
      </td>
      <td align="center" valign="middle"><input name="Add" type="Button" id="Add" value=" &gt;&gt; " onClick="javascript: ClickMe('Add',0);"></td>
      <td valign="top">
       <%
       	sql="SELECT NewsId,Title from News where eventId=" & EventId & " ORDER By Newsid desc"
       	rs.open sql,con,3
       	if not rs.eof then
       %>
        <table width="100%" border="0" cellspacing="0" cellpadding="2">
         <%Do while not rs.eof%>
          <tr>
            <td width="99%"><div align="justify"><font size="2" face="Arial, Helvetica, sans-serif"><li><%=rs("Title")%></li></font></div></td>
            <td width="1%" align="center" valign="top"><a href="javascript: ClickMe('Remove',<%=rs("NewsId")%>);"><img src="../images/icon_closed_topic.gif" alt="Remove" width="15" height="15" border="0" align="absmiddle" vspace="4"></a></td>
          </tr>
         <%rs.movenext
         Loop
         rs.close
         set rs=nothing%>
        </table>
       <%end if 'if not rs.eof then
       %>
       </td>
    </tr>
    <tr>
      <td colspan="3" align="center" valign="middle"><br><br><font size="2" face="Arial, Helvetica, sans-serif"><a href="javascript: window.close();">&#272;&#243;ng c&#7917;a s&#7893;</a></font></td>
    </tr>
  </table>
  <input type="Hidden" name="Action" value="">
  <input type="Hidden" name="R_NewsId" value="">
  <input type="Hidden" name="page" value="">
  <Script language="Javascript">
  	function ClickMe(thevalue,thevalue2)
  	{
  		window.fEvent.Action.value=thevalue;
  		window.fEvent.R_NewsId.value=thevalue2;
  		window.fEvent.submit();
  	}
  	function fSearchOnSubmit(thisvalue)
		{
			window.fEvent.page.value=thisvalue;
			window.fEvent.submit();
		}
  </script>
 </form>
</body>
</html>
