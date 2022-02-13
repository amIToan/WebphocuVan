<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Call Authenticate("None")
%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->

<%Dim rs
Set rs = Server.CreateObject("ADODB.Recordset")
if request.Querystring("delete")="true" then
	if isnumeric(Request.QueryString("filesid")) then
		files_id=Cint(Request.QueryString("filesid"))
		sql="delete files where files_id=" & files_id
		rs.open sql,con,1
	end if
elseif request.Querystring("upload")="true" then
	sError=""
	
	Set Upload = Server.CreateObject("Persits.Upload")
	Upload.codepage=65001
	Upload.SetMaxSize 10000000, True
	'Dung luong Upload 10MB
	Upload.Save
	
	set smallpicture = Upload.Files("smallpicture")
	if smallpicture Is Nothing then
		sError="Bạn không Upload file nào"
	else
	   Filetype = Right(smallpicture.Filename,len(smallpicture.Filename)-Instr(smallpicture.Filename,"."))
	   if Lcase(Filetype)<>"doc" and Lcase(Filetype)<>"rar" and Lcase(Filetype)<>"zip" and Lcase(Filetype)<>"pdf" then
	   		sError="Chỉ Upload được loại file doc, rar, zip, pdf"
	   else
	   		files_id=GetMaxId("files","files_id","")
	   		smallfilename="files_" & files_id & "." & Filetype
			files_size=smallpicture.Size
			files_name=smallpicture.Filename
	   end if
	End If

	if sError="" then
		smallpicture.SaveAs Path & "\" & smallfilename
		if largefilename<>"" then
			largepicture.SaveAs Path & "\" & largefilename
		end if
		sql="insert into files (files_id,files_URL,files_name,files_size) values "
		sql=sql & "(" & files_id
		sql=sql & ",'" & smallfilename & "'"
		sql=sql & ",'" & files_name & "'"
		sql=sql & "," & files_size & ")"
		rs.open sql,con,1
		'response.Write("<script language=""JavaScript"">")
		'response.Write "window.opener.document.forms['" & Request.QueryString("FormName") & "']['" & Request.QueryString("FieldName") & "'].value = '" & NewsImagePath & Filename & "';" & vbNewline &_
		'				"self.close();"
		'response.Write("/script>")
		'response.End()
	end if
	set Upload=nothing
End if
%>
<html>
<head>
<title>image addnew</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<Script language="JavaScript">
	function returnvalue(thisvalue,namevalue)
	{
		opener.InsertNewFile("<a href=\"" + thisvalue + "\">" + namevalue +"</a>");
		window.close();
	}
</Script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form action="<%=Request.ServerVariables("SCRIPT_NAME")%>?upload=true" method="post" enctype="multipart/form-data" name="fAddImage">
  <table width="100%" border="0" cellspacing="5" cellpadding="2">
    <tr align="center" valign="middle"> 
      <td colspan="3"><font size="3" face="Verdana, Arial, Helvetica, sans-serif"><font size="2"> 
        <strong>Upload file</strong> 
        <%if sError<>"" then
			response.Write("<br><font size=""1"">(<font color=""red""><b>*</b></font>&nbsp;" & sError & ")</font>")
		end if%>
        </font></font></td>
    </tr>
    <tr> 
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">File</font></td>
      <td> <input name="smallpicture" type="file" id="largepicture2" size="20">
        <font color="#FF0000" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font> 
        <br>
        <font size="2" face="Arial, Helvetica, sans-serif">*.doc,*.rar,*.zip,*.pdf; 
        &lt;10MB </font></td>
      <td width="30%" align="left"><input type="submit" name="submit" value="Upload files"></td>
    </tr>
    <tr> 
      <td colspan="3"><hr align="center" width="80%" size="1" noshade></td>
    </tr>
    <tr> 
      <td colspan="3"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
          Các file đã Upload</strong></font></div></td>
    </tr>
    <tr> 
      <td colspan="3"> <table width="100%" border="1" cellspacing="0" cellpadding="2">
          <tr> 
            <td width="5%" align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">TT</font></td>
            <td width="35%"> <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
                Tên file</font></div></td>
            <td width="25%" align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Kích 
              cỡ</font></td>
            <td width="25%" align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Ngày 
              Upload </font></td>
            <td width="10%"> <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Xử 
                lý</font></div></td>
          </tr>
          <%
if isnumeric(request.QueryString("page")) and request.QueryString("page")<>0 then
	page=Cint(request.QueryString("page"))
else
	page=1
end if

sql="select * from files order by files_id desc"
rs.PageSize = 10
rs.Open sql, con, 1
if not rs.eof then
	rs.AbsolutePage = CLng(page)
	i=0
	stt=(page-1)* rs.pageSize + 1
		Do while not rs.eof and i<rs.pagesize
%>
          <tr> 
            <td align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=stt%>.</font></td>
            <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
				<a href="javascript: returnvalue('<%=NewsImagePath%><%=rs("files_URL")%>','Download');"> 
					<%=rs("files_name")%></a>
			</font></td>
            <td align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
              <%=rs("files_size")%> bytes</font></td>
            <td align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
	              <%=Day(rs("files_createdate"))%>/<%=Month(rs("files_createdate"))%>/<%=Year(rs("files_createdate"))%>
              </font></td>
            <td align="center"> <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?delete=true&filesid=<%=rs("files_id")%>"> 
              <img src="../imgedit/recycle_bin.jpg" width="15" height="16" border="0" title="  X&#243;a  "></a> 
            </td>
          </tr>
          <%i=i+1
		  	stt=stt+1
		  	rs.movenext
			Loop
		  %>
          <tr> 
            <td colspan="5" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%
					for i=1 to rs.pagecount
					  if i<>page then
						response.Write "<a href=""" & request.ServerVariables("SCRIPT_NAME") & "?page=" & i & """>" & i & "</a>&nbsp;|&nbsp;"
					  else
					  	response.Write "<font color=""red""><b>" & i & "</b></font>&nbsp;|&nbsp;"
					  end if
					Next
					rs.close
				%>
              </font> </td>
          </tr>
          <%End if%>
        </table></td>
    </tr>
  </table>
</form>
</body>
</html>
<%
Set rs = Nothing
con.close
set con=nothing%>
