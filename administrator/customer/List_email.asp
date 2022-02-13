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
	iTim =GetNumeric(Request.form("cbAll"),0)	
	strDieuKien	=	Request.Form("txtDieuKien")
	iDieuKien	=	GetNumeric(Request.Form("selDieuKien"),0)
	iSapXep			=	GetNumeric(Request.Form("selSapXep"),0)
	iTangPricem		=	GetNumeric(Request.Form("raTangPricem"),0)
	iCheckNhom		=		GetNumeric(Request.Form("CheckGroupEmail"),0)
	
	NameProvince		=	Trim(Request.Form("selTinh1"))
	fTien1	=	Chuan_money(Request.Form("txtTien1"))
	cbAll	=	GetNumeric(Request.Form("cbAll"),0)
	m_day	=	GetNumeric(Request.Form("m_day"),0)
	stt		=	GetNumeric(Request.Form("stt"),0)
	
	iTop	=	GetNumeric(Request.Form("txtTop"),0)
	iTopBegin	=	GetNumeric(Request.Form("txtTopBegin"),0)
	
	iCheckThuTay	=	GetNumeric(Request.Form("iCheckThuTay"),0)
	act		=	"SetEmailSend.asp"
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
	act		=	"List_email.asp"
END IF
%>
<html>
	<head>
	<title><%=PAGE_TITLE%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<SCRIPT language=JavaScript1.2 src="/administrator/inc/common.js"></SCRIPT>
	<link href="../../css/styles.css" rel="stylesheet" type="text/css">
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
        <br>
        SEARCH AND SEND
        <br>
    </td>
  </tr>
  <tr>
    <td >
	<table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" style="border:#CCCCCC solid 1px;">
      <tr>
        <td align="center">
		<table width="80%" align="center" cellpadding="2" cellspacing="2" class="CTxtContent">
        <tr>
          <td colspan="4" align="right" valign="middle" ><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Time search:</strong></font>
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
           &nbsp;&nbsp;&nbsp;<br>
			<input name="m_day" type="radio" value="0">
          Date of birth
          <input name="m_day" type="radio" value="1" checked>
          Date input data </div></td>
          </tr>
        
        <tr>
          <td width="15%" align="center" valign="middle" class="CTxtContent" ><div align="right">Input text:</div></td>
          <td width="31%" align="center" valign="middle" class="CTxtContent" ><div align="left"><input name="txtDieuKien" type="text" id="txtDieuKien" value="<%=strDieuKien%>"></div></td>
          <td width="20%" align="center" valign="middle" class="CTxtContent" ><div align="right">Chọn:</div></td>
          <td width="34%" align="center" valign="middle" class="CTxtContent" ><div align="left">
            <select name="selDieuKien" id="selDieuKien">
              <option value="0" selected <%if iDieuKien = 0 then%>selected<%end if%>></option>
              <option value="1" <%if iDieuKien = 1 then%>selected<%end if%>>Name</option>
              <option value="2" <%if iDieuKien = 2 then%>selected<%end if%>>Email</option>
            </select>
          </div></td>
        </tr>
        <tr>
          <td align="right" valign="middle" class="CTxtContent" >
            
		  <input name="CheckGroupEmail" type="hidden" id="CheckGroupEmail" value="1">		  </td>
          <td align="left" valign="middle" class="CTxtContent" ></td>
          <td align="right" valign="middle" class="CTxtContent" ></td>
          <td align="left" valign="middle" class="CTxtContent" ></td>
        </tr>
         <tr>
          <td colspan="4" align="center" valign="middle"  >
		  <span id="ShowGroupEmail" style="display:none">
  		  <table width="100%" border="0" cellspacing="0" cellpadding="0" style="border:#CCCCCC solid 1px;">
			<tr>
            <%
  	sql = "Select * From EmailNhom ORDER BY CatGroup "
  	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open sql,Con,3
	iCol	=	0
	stt=0
	do while not rs.eof 
		stt	=	stt	+ 1
		if iCol >= 2 then
			Response.Write("</tr><tr>")
			iCol = 0
		end if
%>		  
				
              <td><input name="IDNhomEmail<%=stt%>" type="checkbox" value="<%=rs("IDNhomEmail")%>"> <%=rs("TenNhom")%></td>
            
        <%		
		iCol	=	iCol + 1	
		rs.movenext
	loop
  set rs = nothing
%>			
</tr>
          </table>
		  <input type="hidden" value="<%=stt%>" name="stt">
		  </span>		  </td>
        </tr>

        <tr>
          <td colspan="4" align="center" valign="middle" >
		  <input name="cbAll" type="checkbox" id="cbAll" value="1" > All email
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
	<table <%if iCheckThuTay = 0 then%> width="95%" <%else Response.Write("width=""500""") end if%> border="0" align="center" cellpadding="1" cellspacing="1">
<%
	FromDate=Thang1 & "/" & Ngay1 & "/" & Nam1
  	ToDate=Thang2 & "/" & Ngay2 & "/" & Nam2
  	
  	FromDate=FormatDatetime(FromDate)
	ToDate=FormatDatetime(ToDate)
		
	sql="select * from Email"
	if cbAll <> 1 then
	if iTim <> 1 then
		select case iDieuKien
			case 0 
				sql = sql & " WHERE "		
			case 1 
				strDieuKien = Ucase(strDieuKien)
				sql=sql & " where {fn UCASE(Ten)} like N'%" & strDieuKien & "%' and "
			case 2
				strDieuKien = Ucase(strDieuKien)
				sql=sql & " where {fn UCASE(Email)} like N'%" & strDieuKien & "%' and "		
					
		end select
	else
		sql = sql & " WHERE "		
	end if
	
	if m_day = 0 then	
		sql=sql+"  (DATEDIFF(dd,NgaySinh,'" & FromDate & "')<= 0) AND (DATEDIFF(dd,NgaySinh,'" & ToDate &"') >= 0) "
	else
		sql=sql+"  (DATEDIFF(dd,CreateDate,'" & FromDate & "')<= 0) AND (DATEDIFF(dd,CreateDate,'" & ToDate &"') >= 0) " 
	end if
	
	if iCheckNhom > 0 then
		ik	=1
		IDNhom	=	GetNumeric(Request.Form("IDNhomEmail"&ik),0)
		sql=sql+" AND ((IDCongViec = '"& IDNhom &"') "
		for ik = 2 to stt
			IDNhom	=	GetNumeric(Request.Form("IDNhomEmail"&ik),0)
			if IDNhom > 0 then
				sql=sql+" or (IDCongViec = '"& IDNhom &"')  "
			end if
		next	
        sql=sql+")"	
	end if

	end if

  Set rs = Server.CreateObject("ADODB.Recordset")
  rs.open sql,Con,3
  length = rs.recordcount-1
%>
        <tr>
	<td width="35" height="19" align="center" bgcolor="#FFFFCC" class="CFontVerdana10" style="<%=setStyleBorder(1,1,1,1)%>">No.</td>
	<%if iCheckThuTay = 0 then%>	
	<td width="49" bgcolor="#FFFFCC" class="CFontVerdana10" style="<%=setStyleBorder(0,1,1,1)%>" align="center">
	    Choice	</td><%end if%>
	<td width="151" bgcolor="#FFFFCC" class="CFontVerdana10" align="center" style="<%=setStyleBorder(0,1,1,1)%>">Name</td>
	<%if iCheckThuTay = 0 then%>
	<td width="308" bgcolor="#FFFFCC" align="center" class="CFontVerdana10" style="<%=setStyleBorder(0,1,1,1)%>">Email</td>
	<td width="183" bgcolor="#FFFFCC" align="center" class="CFontVerdana10" style="<%=setStyleBorder(0,1,1,1)%>">Group</td>
	<td width="34" align="center" bgcolor="#FFFFCC" class="CFontVerdana10" style="<%=setStyleBorder(0,1,1,1)%>">&nbsp;</td>
	<%end if%>
</tr>
        <%
            i=0
  Do while not rs.eof
		ID		=	rs("ID")
		Ten		=	rs("ten")
		Tel		=	rs("DienThoai")
		Email	=	rs("Email")
		Diachi	=	rs("Diachi")
%>
  		<tr <%if i mod 2=0 and iCheckThuTay = 0  then response.Write("bgcolor=""#E6E8E9""") else response.Write("bgcolor=""#FFFFFF""")%>>
			<td align="center" style="<%=setStyleBorder(1,1,0,1)%>"><%=i+iTopBegin+1%></td>
			<%if iCheckThuTay = 0 then%>
			<td align="center" style="<%=setStyleBorder(1,1,0,1)%>">
			<input type="checkbox" name="CbEmailKhach<%=i%>" value="1">	<br>		
			<input type="hidden" name="IDEmail<%=i%>" value="<%=ID%>">			</td>
			
			<%end if%>
			<td class="CTxtContent" style="<%=setStyleBorder(0,1,0,1)%>">
			<b><%=Ten%></b>
			<br>
			<font class="CSubTitle"><u>Tel</u>:<%=Tel%></font>

			</td>
			
			<td style="<%=setStyleBorder(0,1,0,1)%>" >
			<font class="CTxtContent"><%=Diachi%></font><br>
			<%=Email%>
			<input type="hidden" name="hEmail<%=i%>" value="<%=Email%>">
			<br>              </td>
			<td style="<%=setStyleBorder(0,1,0,1)%>">
				<%=Get_Name_Group(Get_IDGroup_FromEmail(Email,"IDCongViec"))%>&nbsp;</td>
			<td style="<%=setStyleBorder(0,1,0,1)%>">
			<a href="updateEmail.asp?addOrEddit=1&ID=<%=ID%>">
			<img src="../../images/icons/article.gif" width="16" height="16" border="0" align="absmiddle"></a>
			
			<img src="../../images/icons/icon_pmdead.gif" border="0" align="absmiddle" onClick="javascript: yn = confirm('Are you sure delete?'); if(yn) {window.location = 'delEmail.asp?ID=<%=ID%>'}" >
			
					</td>
			
		</tr>    
          <%		

		i=i+1
  	rs.movenext
  loop
%>
	<tr>
		<td>&nbsp;</td>
		<td style="<%=setStyleBorder(1,1,0,1)%>" align="center">
		<%
		iSoEmail	=	i
		%>
		<input type="hidden" name="iSoEmail" value="<%=iSoEmail%>">
		<input type="checkbox" name="CbAllEmail" value="1" onClick="javascript:OnCheckAll()" >		</td>
		<td colspan="2"><input type="submit" name="Submit3" value=" Send email with choice " onClick="javascript: ReSub();"></td>
		<td>&nbsp;</td>
		<td align="center">		</td>			
	</tr>
	<tr>
		<td colspan="6" align="center"><input type="button" name="Submit32" value=" Send all email " onClick="    javascript: window.location = 'SetEmailSend.asp?All=ok'">	</td>
	</tr>
	</table>
<%end if%>
	<br></td>
  </tr>
  <tr>
    <td ></td>
  </tr>
</table>
</form>

</body>
</html>
<script language="javascript">
function OnCheckAll()
{
	iNumSP	=	document.fEmail.iSoEmail.value-1;	
	if (document.fEmail.CbAllEmail.checked == true)
		iCbAll = 1
	else
		iCbAll = 0
	
	for(jj=0;jj<=iNumSP;jj++)
	{
		if (iCbAll == 1)
			str = "document.fEmail.CbEmailKhach"+jj+".checked = true";
		else
			str = "document.fEmail.CbEmailKhach"+jj+".checked = false";
		eval(str);	
	}
	return;
	
}

function ShowEmailGroup()
{
    if (document.fEmail.CheckGroupEmail1.value == "Show email group") {
        document.fEmail.CheckGroupEmail1.value = "Hide email group"
        document.getElementById("ShowGroupEmail").style.display = "";
        document.fEmail.CheckGroupEmail.value = 1
    }
    else {
        document.getElementById("ShowGroupEmail").style.display = "none";
        document.fEmail.CheckGroupEmail1.value = "Show email group"
        document.fEmail.CheckGroupEmail.value = 0

    }


}
</script>

