<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_Ads.asp" -->
<%
f_permission = administrator(false,session("user"),"m_ads")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
	Call AuthenticateWithRole(AdvertisementCategoryId,Session("LstRole"),"ap")
	Ads_id=GetNumeric(Request.querystring("id"),0)
	
	CatNum_Selected=CLng(Request.Form("CatId").Count)
	Ads_OnlineChildren_Selected=CLng(Request.Form("Ads_OnlineChildren").Count)
	if Request.Form("action")="Insert" and CatNum_Selected>0 then
		Ads_Order=GetMaxId("AdsDistribution", "Ads_Order", "")
		sql="delete AdsDistribution where Ads_id=" & Ads_id
		For i = 1 to CatNum_Selected
			sql=sql & ";insert into AdsDistribution (Ads_id,CategoryId,Ads_OnlineChildren,Ads_Order) values "
			sql=sql & "(" & Ads_id
			sql=sql & "," & Clng(Request.Form("CatId")(i))
			for j=1 to Ads_OnlineChildren_Selected
				if Clng(Request.Form("CatId")(i))=CLng(Request.Form("Ads_OnlineChildren")(j)) then
					sql=sql & ",1"
					Exit for
				end if
			next
			'Response.write "Ads_OnlineChildren_Selected: " & Ads_OnlineChildren_Selected & ",j=" & j
			if j>Ads_OnlineChildren_Selected then
				sql=sql & ",0"
			end if
			sql=sql & "," & Ads_Order + i -1 & ")"
		Next
		Dim rs
		Set rs=Server.CreateObject("ADODB.Recordset")
		rs.open sql,con,1
		set rs=nothing
		con.close
		set con=nothing
		response.Write "<script language=""JavaScript"">" & vbNewline &_
			"<!--" & vbNewline &_
				"window.opener.location.reload();" & vbNewline &_
				"window.close();" & vbNewline &_
			"//-->" & vbNewline &_
			"</script>" & vbNewline
		response.end
	end if
	Dim rsCat
	Set rsCat=Server.CreateObject("ADODB.Recordset")
	
	sql="SELECT DISTINCT CategoryId, Ads_OnlineChildren " &_
		"FROM	AdsDistribution ad " &_
		"WHERE     (Ads_id = " & Ads_id & ")"
	rsCat.open sql,con,3
	LstCat=""
	LstOnlineChildren=""
	Do while not rsCat.eof
		LstCat=LstCat & Cstr(rsCat("CategoryId")) & " "
		LstOnlineChildren=LstOnlineChildren & Cstr(rsCat("CategoryId")) & Cstr(rsCat("Ads_OnlineChildren")) & " "
	rsCat.movenext
	Loop
	rsCat.close
	LstCat=Trim(LstCat)
	LstOnlineChildren=Trim(LstOnlineChildren)
	
	sql="SELECT	CategoryID, CategoryName, CategoryLevel " &_
		"FROM   NewsCategory " &_
		"WHERE  CategoryStatus = 1 OR CategoryStatus = 2 " &_
		"ORDER BY LanguageId DESC, CategoryOrder"
		
	
	rsCat.open sql,con,3
%>
<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="fMultiChoice" method="post" action="<%=Request.ServerVariables("Script_name")%>?id=<%=Ads_id%>">
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr align="center"> 
    <td height="28" valign="middle"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Chọn 
      các chuyên mục hiển thị quảng cáo</strong></font></td>
  </tr>
  <tr> 
    <td><table width="100%" border="0" cellpadding="2" cellspacing="1" bordercolor="#000000" bgcolor="#000000">
        <tr align="center" bgcolor="#FFFFFF"> 
            <td><font size="2" face="Arial, Helvetica, sans-serif">Danh sách chuyên mục</font></td>
            <td align="center"><img src="../images/icon_folder_unlocked.gif" width="15" height="15" border="0" align="absmiddle" alt="Hiển thị cả ở các chuyên mục con"></td>
        </tr>
        <%i=0
        Do while not rsCat.eof
        	i=i+1%>
        <tr bgcolor="<%if i mod 2 =0 then%>#FFFFFF<%else%>#E6E8E9<%end if%>"> 
          <td><font size="2" face="Arial, Helvetica, sans-serif">
          	<%for j=2 to Clng(rsCat("CategoryLevel"))
        		response.write "&nbsp;&nbsp;&nbsp;&nbsp;"
          	Next%>
          	<input type="checkbox" name="CatId" value="<%=rsCat("CategoryId")%>"<%if CheckCatInLstCat(LstCat,rsCat("CategoryId")) then%> checked<%end if%>>
          	<%if Clng(rsCat("CategoryLevel"))=1 then
          		response.write "<strong>" & rsCat("CategoryName") & "</strong>"
          	else
          		response.write rsCat("CategoryName")
          	end if%>
          </font></td>
          <td align="center"><input type="checkbox" name="Ads_OnlineChildren" value="<%=rsCat("CategoryId")%>"<%if CheckOnlineChildrenInList(rsCat("CategoryId"),LstOnlineChildren) then%> checked<%end if%>></td>
        </tr>
        <%rsCat.moveNext
        Loop
        rsCat.close
        set rsCat=nothing
        con.close
        set con=nothing
        %>
      </table></td>
  </tr>
  <tr>
      <td align="center">
      	<input type="submit" name="Submit" value=" Chọn ">
      	<input type="hidden" name="action" value="Insert">
      </td>
  </tr>
</table></form>
</body>
</html>
