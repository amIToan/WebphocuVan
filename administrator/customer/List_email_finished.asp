<%session.CodePage=65001%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<%
f_permission = administrator(false,session("user"),"m_customer")
if f_permission < 3 then
	response.Redirect("/administrator/info.asp")
end if
%>

<%IF Request.form("action")="Search" then
	Ngay1=GetNumeric(Request.form("Ngay1"),0)
	Thang1=GetNumeric(Request.form("Thang1"),0)
	Nam1=GetNumeric(Request.form("Nam1"),0)
	Ngay2=GetNumeric(Request.form("Ngay2"),0)
	Thang2=GetNumeric(Request.form("Thang2"),0)
	Nam2=GetNumeric(Request.form("Nam2"),0)
	strDieuKien	=	Request.Form("txtDieuKien")
    iDieuKien	=	GetNumeric(Request.Form("selDieuKien"),0)
	cbAll	=	GetNumeric(Request.Form("cbAll"),0)
	act		=	"List_email_finished.asp"
ELSE
	Day1 = now() - 30
	Ngay1=Day(Day1)
	Thang1=Month(Day1)
	Nam1=Year(Day1)
	Ngay2=Day(now())
	Thang2=Month(now())
	Nam2=Year(now())
	iDieuKien = 0
	cbAll	=	0
	act		=	"List_email_finished.asp"
END IF
%>
<html>
	<head>
	<title><%=PAGE_TITLE%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
	<link href="../../css/styles.css" rel="stylesheet" type="text/css">
	    <style type="text/css">
            .auto-style1 {
                width: 758px;
            }
            .auto-style2 {
                height: 22px;
            }
        </style>
	</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
	img	="../../images/icons/new_mail_accept.gif"
	Title_This_Page="Khách hàng -> Danh sách email khách hàng"
	Call header()
	
	
%>

<form name="fEmail" method="post" action="<%=act%>" >
<%IF Request.form("action")<>"Search" then%>
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" class="CTieuDe" >
        HISTORY</td>
  </tr>
  <tr>
    <td >
	<table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" style="border:#CCCCCC solid 1px;">
      <tr>
        <td align="center" class="auto-style1">
		<table width="80%" align="center" cellpadding="2" cellspacing="2" class="CTxtContent">
        <tr>
          <td colspan="2" align="right" valign="middle" ><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Time search:</strong></font>
            <%
			Call List_Month_WithName(Thang1,"MM","Thang1")
            Call List_Date_WithName(Ngay1,"DD","Ngay1")
			Call  List_Year_WithName(Nam1,"YYYY",1960,"Nam1")
		%>
            <img src="../images/right.jpg" width="9" height="9" align="absmiddle">
            <%
            Call List_Month_WithName(Thang2,"MM","Thang2")
			Call List_Date_WithName(Ngay2,"DD","Ngay2")
			Call  List_Year_WithName(Nam2,"YYYY",1960,"Nam2")
		%>
</td>
          </tr>
        
        <tr>
          <td width="15%" align="center" valign="middle" class="CTxtContent" ><div align="right">Input email:</div></td>
          <td width="31%" align="center" valign="middle" class="CTxtContent" ><div align="left"><input name="txtDieuKien" type="text" id="txtDieuKien" value="<%=strDieuKien%>"></div></td>
          <td width="20%" align="center" valign="middle" class="CTxtContent" ><div align="right">choice:</div></td>
          <td width="34%" align="center" valign="middle" class="CTxtContent" ><div align="left">
            <select name="selDieuKien" id="selDieuKien">
              <option value="0" selected <%if iDieuKien = 0 then%>selected<%end if%>></option>
              <option value="1" <%if iDieuKien = 1 then%>selected<%end if%>>From</option>
              <option value="2" <%if iDieuKien = 2 then%>selected<%end if%>>To</option>
            </select>
          </div></td>
        </tr>

        <tr>
          <td colspan="2" align="center" valign="middle" >
            <input type="hidden" name="action" value="Search">
            <input type="submit" name="Submit11" value="       SEARCH     " >          </td>
          </tr>
      </table>		</td>
      </tr>
    </table>
<%end if%>	
	<br>
	<br>
<%IF Request.form("action")="Search" then

%>
	<table width="95%"  border="0" align="center" cellpadding="1" cellspacing="1">
        <tr class="CTieuDeNho" >
        <td style="<%=setStyleBorder(1,1,1,1)%>" align="center" class="auto-style2">
	        From</td>
        <td style="<%=setStyleBorder(0,1,1,1)%>"  align="center" class="auto-style2">
	        To</td>
        <td style="<%=setStyleBorder(0,1,1,1)%>" align="center" class="auto-style2">
	        Title</td>
        <td style="<%=setStyleBorder(0,1,1,1)%>" align="center" class="auto-style2">
	        Date</td>
        <td style="<%=setStyleBorder(0,1,1,1)%>" align="center" class="auto-style2">
	        Tool</td>
        </tr>
<%
	FromDate=Thang1 & "/" & Ngay1 & "/" & Nam1
  	ToDate=Thang2 & "/" & Ngay2 & "/" & Nam2
  	
  	FromDate=FormatDatetime(FromDate)
	ToDate=FormatDatetime(ToDate)
		
	sql="select * from EmailCompose"
	if iTim <> 1 then
		select case iDieuKien
			case 0 
				sql = sql & " WHERE "		
			case 1 
				strDieuKien = Ucase(strDieuKien)
				sql=sql & " where {fn UCASE(FromEmail)} like N'%" & strDieuKien & "%' and "
			case 2
				strDieuKien = Ucase(strDieuKien)
				sql=sql & " where {fn UCASE(ToEmail)} like N'%" & strDieuKien & "%' and "		
		end select
	else
		sql = sql & " WHERE "		
	end if
	sql=sql+"  (DATEDIFF(dd,CreateDate,'" & FromDate & "')<= 0) AND (DATEDIFF(dd,CreateDate,'" & ToDate &"') >= 0) " 
	
    sql=sql & " order by CreateDate desc"

    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.open sql,Con,3

 
  Do while not rs.eof
        ID             =   rs("ID")
		FromEmail		=	rs("FromEmail")
		ToEmail		    =	rs("ToEmail")
		TitleEmail		=	rs("TitleEmail")
		CreateDate	    =	rs("CreateDate")
%>
       <tr>
			<td class="CTxtContent" style="<%=setStyleBorder(1,1,0,1)%>">
                <%= FromEmail%>	
			</td>
			<td style="<%=setStyleBorder(0,1,0,1)%>" >
			<%= ToEmail%>	
                 </td>
			<td style="<%=setStyleBorder(0,1,0,1)%>">
				<%=TitleEmail%>
			</td>
			<td style="<%=setStyleBorder(0,1,0,1)%>">
                <%=CreateDate%>
			</td>
			<td style="<%=setStyleBorder(0,1,0,1)%>" align="center">
                <a href="#" onClick="javascript: yn = confirm('Do you sure delete?'); if(yn) {window.location = 'DelMailHistory.asp?ID=<%=ID%>'}">Delete</a>
            </td>           
		</tr>
<%        		
  	rs.movenext
  loop
end if
 %>
	
       <tr>
			<td>
               
			</td>
			<td >
	
                 </td>
			<td >
			
			</td>
			<td>
              
			</td>
			<td style="<%=setStyleBorder(1,1,1,1)%>">
                <a href="#" onClick="javascript: yn = confirm('Do you sure delete all history of send list?'); if(yn) {window.location = 'DelMailHistory.asp?Delall=all'}">Empty history</a>
                
			</td>           
		</tr>
	</table>

</form>

</body>
</html>


