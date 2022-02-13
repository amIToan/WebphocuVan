 <%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->

<!--#include virtual="/administrator/inc/func_DateTime.asp" -->
<!--#include virtual="/include/func_calculate.asp" -->
<!--#include virtual="/administrator/inc/func_tiny.asp" -->
<html>
<head>
<title><%=PAGE_TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="../../css/styles.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
	AcUpdate = GetNumeric(Request.Form("AcUpdate"),0)
	if AcUpdate = 0 then
		AcUpdate	=	GetNumeric(Request.QueryString("AcUpdate"),0)
	end if
	if AcUpdate = 1 or AcUpdate = 2 then
		Gio1	= 	GetNumeric(Request.Form("Gio1"),0)
		phut1	= 	GetNumeric(Request.Form("phut1"),0)
		Gio2	= 	GetNumeric(Request.Form("Gio2"),0)
		phut2	= 	GetNumeric(Request.Form("phut2"),0)
	
		Ngay1	=	Request.form("Ngay1")
		Thang1	=	Request.form("Thang1")
		Nam1	=	Request.form("Nam1")
		Ngay2	=	Request.form("Ngay2")
		Thang2	=	Request.form("Thang2")
		Nam2	=	Request.form("Nam2")	
		
		iKichHoat	= GetNumeric(Request.Form("iKichHoat"),0)
				
		Gio11	= 	Cstr(Gio1) + ":" + Cstr(Phut1)
		Gio22	= 	Cstr(Gio2) + ":" + Cstr(Phut2)
		Ngay11	=	Thang1+"/"+Ngay1+"/"+Nam1
		Ngay22	=	Thang2+"/"+Ngay2+"/"+Nam2
		FromDate	=	Ngay11 + " "	+ Gio11
		ToDate	=	Ngay22 + " "	+ Gio22	
		
		iType	=	GetNumeric(Request.form("selType"),0)
		iSubjects	=	GetNumeric(Request.form("selSubjects"),0)
	
		TieuDe	=	Replace(Trim(Request.Form("TieuDe")),"'","''")
		TieuDe	=	Replace(TieuDe,chr(13) & chr(10),"<br>")
		
		NoiDung	=	Replace(Trim(Request.Form("NoiDung")),"'","''")
		NoiDung	=	Replace(NoiDung,chr(13) & chr(10),"<br>")

		iKichHoat	=	GetNumeric(Request.Form("cbShow"),0)	
	end iF
	select case AcUpdate
	case 1
		ID		= GetNumeric(Request.Form("ID"),0)
		sql		= "update ThongBao set "
		sql		= sql+ " TieuDe = N'"& TieuDe &"',"
		sql		= sql+" NoiDung = N'"& NoiDung &"',"
		sql		= sql+" FromDate = '"& FromDate &"',"
		sql		= sql+" ToDate = '"& ToDate &"',"
		sql		= sql+" Type = '"& iType &"',"
		sql		= sql+" Subjects = '"& iSubjects &"',"		
		sql		= sql+" Kichhoat = '"& iKichHoat &"'"
		sql		= sql+" where ID ="&ID	
		isBack	=	2		
	case 2
		sql 	=	"insert into ThongBao(TieuDe,NoiDung,FromDate,ToDate,Type,Subjects,KichHoat) values(N'"&TieuDe &"',N'"& NoiDung&"','"& FromDate&"','"& ToDate&"','"&iType&"','"& iSubjects &"','"& iKichHoat &"')"	
		isBack	=	2	
	case 3
		ID		= GetNumeric(Request.QueryString("ID"),0)
		sql		=	"Delete ThongBao where ID="&ID
		isBack	=	1
	end select
	if AcUpdate > 0 then
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.open sql,con,1
		
		set rs=nothing
		response.Write "<script language=""JavaScript"">" & vbNewline &_
		"<!--" & vbNewline &_
			"alert('Đã cập nhận xong thưa ông chủ! Mời ngài nhấn F5 để thay đổi được cập nhật');" & vbNewline &_
			"history.back;" & vbNewline &_
			"history.back;" & vbNewline &_
			"window.reload;" & vbNewline &_
		"//-->" & vbNewline &_
		"</script>" & vbNewline				
	end if
	
%>
<%
	
	img	=	"../../images/icons/announce.gif"
	Title_This_Page="Quản lý ->Thông báo nội bộ"
	Call header()
	
	
%>

  <table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td ></td>
    </tr>
	<tr>
		<td >
<%
	isEdit = GetNumeric(Request.QueryString("isEdit"),0)
	select Case isEdit
	case 0
		call List_ThongBao()
	case 1
		id = GetNumeric(Request.QueryString("ID"),0)
		Call CTietThongBao(ID)
	case 2
		 Call CTietThongBao(0)
	end select
	
%>			
		</td>
	</tr>
  <td ></td>
  </tr>
  </table>
<%Call Footer()%>

<%
sub CTietThongBao(ID)
ID	=	GetNumeric(ID,0)
if ID <> 0 then
	sql = "select * from ThongBao where ID ="&id
	set rsTemp  = Server.CreateObject("ADODB.recordset")
	rsTemp.open sql,con,1
	if not rsTemp.eof then
		FromDate	=	rsTemp("FromDate")
		ToDate		=	rsTemp("ToDate")
		if isdate(ToDate)=false then
			ToDate	=	 Now
		end iF	
		TieuDe		=	trim(rsTemp("TieuDe"))
		NoiDung		=	Trim(rsTemp("NoiDung"))
		iKichHoat	=	rsTemp("KichHoat")
		isType		=	rsTemp("Type")
		isSubjects	=	rsTemp("Subjects")
	end if
else
	
	FromDate	=	Now
	ToDate		=	now + 1	
	TieuDe		=	""
	NoiDung		=	""
end if
	Gio1 	= 	Hour(FromDate)
	Gio2	=	Hour(ToDate)
	Phut1	=	MINUTE(FromDate)
	Phut2	=	MINUTE(ToDate)

%>
<FORM action="upThongBao.asp" method="post" name="fThongBao">
  <table width="80%" border="0" align="center" cellpadding="2" cellspacing="2" class="CTxtContent">
	<tr>
	  <td >
	    <table width="100%" border="0" cellpadding="2" cellspacing="2" class="CTxtContent">
          <tr>
            <td width="15%">Bắt đầu:</td>
            <td width="14%"><select name="Gio1">
              <%for i = 0 to 23%>
              <option value="<%=i%>" <%if Gio1 = i then%>selected<%end if%>><%=i%></option>
              <%next%>
            </select>
h
<select name="Phut1">
  <%for i = 0 to 50 step 10%>
  <option value="<%=i%>" <%if Phut1 = i then%>selected<%end if%>><%=i%></option>
  <%next%>
</select>
p &nbsp;</td>
            <td width="43%">ngày
            <%
			Call List_Date_WithName(Day(FromDate),"DD","Ngay1") 
			Call List_Month_WithName(month(FromDate),"MM","Thang1")
			Call  List_Year_WithName(year(FromDate),"YYYY",2004,"Nam1")%></td>
            <td width="28%" rowspan="3"><input name="cbShow" type="checkbox" id="cbShow" value="1"  <%if iKichHoat <> 0 then%>  checked="checked" <%end if%>>
Kích hoạt</td>
          </tr>
          <tr>
            <td>Kết thúc:</td>
            <td><select name="Gio2" >
              <%for i = 0 to 23%>
              <option value="<%=i%>" <%if Gio2 = i then%>selected<%end if%>><%=i%></option>
              <%next%>
            </select>
h
<select name="Phut2">
  <%for i = 0 to 50 step 10%>
  <option value="<%=i%>" <%if Phut2 = i then%>selected<%end if%>><%=i%></option>
  <%next%>
</select>
p </td>
            <td>ngày
            <%
			Call List_Date_WithName(Day(ToDate),"DD","Ngay2") 
			Call List_Month_WithName(month(ToDate),"MM","Thang2")
			Call  List_Year_WithName(year(ToDate),"YYYY",2004,"Nam2")%></td>
          </tr>
          <tr>
            <td>Loại thông báo:</td>
            <td><select name="selType">
              <option value="0" selected <%if isType = 0 then Response.Write("selected") end if%>>Thông thường</option>
              <option value="1" <%if isType = 1 then Response.Write("selected") end if%>>Thông báo nghỉ</option>
            </select></td>
            <td>Đối tượng:
              <select name="selSubjects">
                <option value="0" selected <%if isSubjects = 0 then Response.Write("selected") end if%> >Nội bộ</option>
                <option value="1" <%if isSubjects = 1 then Response.Write("selected") end if%>>Khách hàng</option>
                <option value="2" <%if isSubjects = 2 then Response.Write("selected") end if%>>Tất cả</option>
              </select></td>
          </tr>

        </table>
</td>
	</tr>

	<tr>
	  <td><div align="center">Tóm tắt:<br> 
	      <textarea name="TieuDe" cols="100" rows="5" id="TieuDe"><%=TieuDe%></textarea>
	  </div></td>
	</tr>
	<tr>
	  <td><div align="center">Nội dung: <br>
	      <textarea name="noidung" cols="100" rows="20" id="noidung"><%=NoiDung%></textarea>
	  </div></td>
	</tr>
	<tr>
	  <td>&nbsp;</td>
    </tr>
	<tr>
	  <td align="center">
	  <input type="hidden" name="AcUpdate" value="<%=isEdit%>">
	  <input type="hidden" name="ID" value="<%=ID%>">
	  <input type="submit" name="Submit" value="  OK  " onClick="javascript checkSub();"></td>
    </tr>
  </table>
</FORM>		
<%end sub%>
<%sub List_ThongBao()%>
	  <table width="80%" border="0" align="center" cellpadding="1" cellspacing="1" class="CTxtContent">
	<tr>
	  <td width="14%" align="center" style="<%=setStyleBorder(1,1,1,1)%>"><strong>Ngày thông báo </strong></td>
	  <td width="78%" align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>	Tóm tắt</strong></td>
	  <td width="8%" align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>TOOLS</strong></td>
	</tr>
	
	<%	
	sql = "Select * from ThongBao order by id Desc"
	set rsTB = Server.CreateObject("ADODB.recordset")
	rsTB.Open sql,con,1
	if not rsTB.eof then
	do while not rsTB.eof
	%>
	<tr>
	  <td align="center" style="<%=setStyleBorder(1,1,0,1)%>"><%=Day(rsTB("FromDate"))%>/<%=month(rsTB("FromDate"))%>/<%=Year(rsTB("FromDate"))%></td>
	  <td style="<%=setStyleBorder(0,1,0,1)%>"><%=rsTB("TieuDe")%></td>
	  <td style="<%=setStyleBorder(0,1,0,1)%>" align="center" >
	  <%
	  if rsTB("Kichhoat") <> 0 then
	  	Response.Write(" <img src=""../images/icon-activate.gif"" width=""16"" height=""16"">")
	  else
		Response.Write(" <img src=""../images/icon-deactivate.gif"" width=""16"" height=""16"">")
	  end if
	  %>	  
	  <a href="upThongBao.asp?isEdit=1&ID=<%=rsTB("ID")%>" class="CSubMenu" target="_parent">
	  <img src="../images/icon_edit_topic.gif" width="15" height="15" border="0" align="absmiddle"> </a> 
	  <img src="../../images/icons/icon_closed_topic.gif" width="15" height="15"  border="0" align="absmiddle" onClick="javascript: yn = confirm('Mời bạn xuy nghĩ kỹ trước khi xóa đó lệnh này rất nguy hiểm? \n Bạn có chắc chắn xóa thông báo này không?'); if(yn) {window.location = 'upThongBao.asp?ID=<%=rsTB("ID")%>&AcUpdate=3'}">	  </td>
	</tr>
	<%
			rsTB.movenext
		loop
  	else
%>	
 	<tr>
	  <td colspan="3" align="center">Hiện không có thông báo nào</td>
	  </tr>
<%	end if
  %>
	<tr>
	  <td align="center"></td>
	  <td><br></td>
	  <td align="center" > <a href="upThongBao.asp?isEdit=2" class="CSubMenu">Thêm</a></td>
	</tr>

  </table>

  	
<%end sub%>
<script src='../inc/news.js'></script>
<script>VISUAL=4; FULLCTRL=1;</script>
<script src='../js/quickbuild.js'></script>
<script>changetoIframeEditor(document.fThongBao.TieuDe)</script>
<script>changetoIframeEditor(document.fThongBao.noidung)</script>
</body>
</html>