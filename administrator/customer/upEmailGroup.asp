<%session.CodePage=65001%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_Donhang.asp" -->
<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<!--#include virtual="/include/func_calculate.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<%
f_permission = administrator(false,session("user"),"m_customer")
if f_permission < 2 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
	isAction	=	GetNumeric(Request.Form("radiobutton"),0)
    IDNhomEmail	=	GetNumeric(Request.Form("selNhomEmail2"),0)
	select case isAction 
		case 1
			TenNhom	=	Trim(Request.Form("tennhom"))
		    sql = "update EmailNhom set "
			sql = sql+"TenNhom = N'"& TenNhom &"'"
			sql	= sql+	" where IDNhomEmail="&IDNhomEmail
		case 2
			iCatGroup	=	GetNumeric(Request.Form("rCatGroup"),0)
			TenNhom	=	Trim(Request.Form("tennhom"))
			sql = "Insert into EmailNhom(CatGroup,TenNhom)"
			sql = sql + "Values('2',N'"& TenNhom &"')"
		case 3
			iCatGroup	=	GetNumeric(Request.Form("rCatGroup"),0)
			sql = "delete EmailNhom where IDNhomEmail="&IDNhomEmail			
	end select
	if (isAction <> 0) then
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.open sql,con,1
		set rs=nothing
	%>
		<script language="javascript">
			window.reload();
		</script>
	<%	
	end if
	
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
	img	="../../images/icons/icon_email1.gif"
	Title_This_Page="Khách hàng -> Soạn và gửi email khách hàng"
	Call header()

%>

	  <form action="upEmailGroup.asp"  name="NhomEmail" method="post">
	    <table width="800px" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td valign="middle"><table width="100%" border="0" align="left" cellpadding="0" cellspacing="0" class="CTxtContent">

              <tr>
                <td colspan="2" align="center" class="CTieuDe">
                    UPDATE EMAIL GROUP
					</td>
              </tr>

              <tr>
                <td style="vertical-align:top;">				
					
<%
					sql = "Select * From EmailNhom where CatGroup=2"
					set rsNEmail = Server.CreateObject("ADODB.recordset")
					rsNEmail.open sql,con,1
					iCount = rsNEmail.recordcount
								
				%>
                  <select name="selNhomEmail2" size="<%=iCount%>" id="selNhomEmail" style="border:#CCCCCC solid 1" onChange="javascript:onSel2();">
                    <%
					do while not rsNEmail.eof 
					%>
                    <option value="<%=rsNEmail("IDNhomEmail")%>"><%=rsNEmail("TenNhom")%></option>
                        <%
						rsNEmail.movenext
					loop
					set rsNEmail = nothing
					%>
                    </select>						
								</td>
                <td style="vertical-align:top;">
                    <table  border="0" cellpadding="2" cellspacing="2" class="CTxtContent" style="border:#CCCCCC solid 1px;">
                <tr>
                  <td >Email group: </td>
                  <td ><input name="tennhom" type="text" class="CTextBoxUnder" id="tennhom" size="50"></td>
                </tr>
                <tr>
                  <td colspan="2"></td>
                </tr>
                <tr>
                  <td colspan="2" align="center">
                      <input name="radiobutton" type="hidden" value="0">
                      <input name="update" type="button" id="update" value=" Update " onClick="javascript: ReSubmitDuyet('1');">
                      <input name="Add" type="button" id="Add" value=" Add " onClick="javascript: ReSubmitDuyet('2');" >
                      <input name="Delete" type="button" id="Delete" value=" Delete " onClick="javascript: ReSubmitDuyet('3');">
                </tr>
              </table></td>
              </tr>
            </table>
            </td>
          </tr>
        </table>
	    </form>

</body>
</html>
<script language="javascript">
function ReSubmitDuyet(ivar)
{
   
	if(document.NhomEmail.tennhom.value =='')
	{	
		alert('Please, input email group')		
		document.NhomEmail.tennhom.focus();
		return;
	}
	document.NhomEmail.radiobutton.value = ivar;
	if (ivar == 3)
	{
	    yn = confirm('Are you sure delete??');
	    if (yn == false)
	        return;
	}
	document.NhomEmail.submit();


}
function onSel()
{
	document.NhomEmail.tennhom.value = document.NhomEmail.selNhomEmail0.options[document.NhomEmail.selNhomEmail0.selectedIndex].text 
}
function onSel1()
{
	document.NhomEmail.tennhom.value = document.NhomEmail.selNhomEmail1.options[document.NhomEmail.selNhomEmail1.selectedIndex].text 
}
function onSel2()
{
	document.NhomEmail.tennhom.value = document.NhomEmail.selNhomEmail2.options[document.NhomEmail.selNhomEmail2.selectedIndex].text 
}

</script>