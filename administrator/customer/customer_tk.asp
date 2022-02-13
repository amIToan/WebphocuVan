<%session.CodePage=65001%>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/include/func_common.asp"-->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<%
f_permission = administrator(false,session("user"),"m_customer")
if f_permission = 0 then
	response.Redirect("/administrator/info.asp")
end if
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Tài khoản cá nhân</title>
<link href="../../css/styles.css" rel="stylesheet" type="text/css"></head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
	CMND	=	Request.QueryString("param")	
	iTotalIdea			=	GetNumeric(fTotalIdea(CMND),0)*250
	iTotalUser			=	GetNumeric(fTotalUser(CMND),0)
	iTotalIni			=	GetNumeric(fIniTKhoan(CMND),0)
	iTotalTK			=	iTotalIdea + iTotalIni
	iTotalElse			=	iTotalTK	-	iTotalUser
	
%>
				  
				  <table width="550" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
                    <tr>
                      <td colspan="4"  background="../../images/TabChinh.gif" height="30" class="CTieuDeNho" style="background-repeat:no-repeat">
					  &nbsp;Tài khoản tăng  </td>
                    </tr>
					<tr>
                      <td align="center" style="<%=setStyleBorder(1,1,0,1)%>" height="28"><strong>Ngày</strong></td>
                      <td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><strong>Lý Do </strong></td>
                      <td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><strong>Số tiền </strong></td>
                      <td align="center" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
					</tr>
					<%
					sql = "SELECT * FROM TaiKhoan where CMND='"& CMND &"'"
					Set rsTK=Server.CreateObject("ADODB.Recordset")
					rsTK.open sql,con,1
					fTongTienTK = 0
				if not rsTK.eof then
					
					do while not rsTK.eof 
						ID 			=	rsTK("ID")
						DatePS 		= 	rsTK("iniDates")
						iTienTK		= 	rsTK("iniTK")
						fTongTienTK	=	fTongTienTK+iTienTK
					%>

                    <tr>
                      <td width="64" align="center" style="<%=setStyleBorder(1,1,0,1)%>"><%=Day(DatePS)%>/<%=Month(DatePS)%>/<%=Year(DatePS)%></td>
                      <td width="329" style="<%=setStyleBorder(0,1,0,1)%>" height="26">&nbsp;<%=rsTK("Lydo")%></td>
                      <td width="108" align="right" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(iTienTK)&DonviPrice%></td>
                      <td width="49" align="center" style="<%=setStyleBorder(0,1,0,1)%>">					  <a href="EditTKCustomer.asp?Status=edit&ID=<%=ID%>" target="_parent"> 
					  <img src="../images/icon_edit_topic.gif" width="15" height="15" border="0" align="absmiddle"> </a>
					  <a href="EditTKCustomer.asp?Status=del&ID=<%=ID%>" target="_blank"> 
					  <img src="../../images/icons/icon_closed_topic.gif" width="15" height="15"  border="0" align="absmiddle">		</a></td>
                    </tr>
					<%
						rsTK.movenext
					loop
					set rsTK = nothing
					%>
					
					<%
					else
					
					%>
                   <tr>
                     <td colspan="4" align="center">Chưa phát sinh</td>
                    </tr>
					<%end if%>	
                    <tr>
                      <td colspan="2" a align="center"><strong>Tổng:</strong></td>
                      <td align="right"><strong><%= Dis_str_money(fTongTienTK)& DonviPrice%></strong></td>
                      <td align="center"> <a href="EditTKCustomer.asp?Status=add&CMND=<%=CMND%>" target="_parent"> 
					  <img src="../images/icon-banner-new.gif" width="16" height="16" border="0">
					  </a></td>
                    </tr>
                    <tr>
                      <td class="CTxtContent" align="center"><img src="images/dot.png" width="3" height="4"></td>
                      <td height="25" class="CTxtContent"><strong><font color="#2C55A5">Thẻ Giảm giá (Vourcher) </font></strong></td>
                      <td class="CTxtContent"><input name="Vourcher" type="text" id="Vourcher" value="<%=Vourcher%>" size="45"  class="CTextBoxUnder" >                   </td>
                      <td class="CTxtContent"><input name="Bt_Vourcher" type="submit" id="Bt_Vourcher" value="&gt;&gt;"></td>
                    </tr>					

                      <tr>
                        <td colspan="4">&nbsp;</td>
                      </tr>
                      <tr>
                      <td colspan="4" background="../../images/TabChinh.gif" height="30" class="CTieuDeNho" style="background-repeat:no-repeat">
					  &nbsp;Số tiền đã sử dụng					  </td>
                    </tr>
                      <tr>
                        <td align="center" style="<%=setStyleBorder(1,1,0,1)%>" height="28"><strong>Ngày đặt </strong></td>
                        <td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><strong>Đơn Hàng </strong></td>
                        <td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><strong>Số tiền </strong></td>
                        <td align="center" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
                      </tr>
					  <% 
					  	sql = "SELECT SanPhamUser.SanPhamUser_ID,CMND, SanPhamUser_Date, TruTrongTaiKhoan FROM  SanPham_pay INNER JOIN SanPhamUser ON SanPham_pay.SanPhamUser_ID = SanPhamUser.SanPhamUser_ID WHERE (SanPhamUser_Status <= 2) and CMND ='"& CMND &"'"
					 	Set rsTK=Server.CreateObject("ADODB.Recordset")
						rsTK.open sql,con,1
						fPSPricemTK = 0
					if not rsTK.eof then
						do while not rsTK.eof 
						DatePSPricem 		= 	rsTK("SanPhamUser_Date")
						iPricemTK		= 	rsTK("TruTrongTaiKhoan")
						fPSPricemTK	=	fPSPricemTK+iPricemTK
						tempSanPhamUser = "ea4bb1a25bfd1b838a8a940d02c8d8ec"&rsTK("SanPhamUser_ID")&"d51dccca677d8049c356f8e5c830d7fc"
					  %>
                    <tr>
                     <td width="64" align="center" style="<%=setStyleBorder(1,1,0,1)%>" height="26"> <%=Day(DatePSPricem)%>/<%=Month(DatePSPricem)%>/<%=Year(DatePSPricem)%></td>
                      <td width="329" style="<%=setStyleBorder(0,1,0,1)%>">
					  <a href="http://xbook.com.vn/ReportDH.asp?xboooooooo000000k=<%=tempSanPhamUser%>" class="CSubMenu" target="_blank">
			<%
				munb=1000+rsTK("SanPhamUser_ID")
				strTemp	="XB"+CStr(munb)
				Response.Write(strTemp)
			%>
					  </a></td>
                      <td width="108" align="right" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(iPricemTK)&DonviPrice%></td>
                      <td width="49" align="right" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
                    </tr>
					<%
						rsTK.movenext
						loop
						set rsTK = nothing
%>
                    <tr>
                      <td colspan="2" align="center"><strong>Tổng:</strong></td>
                      <td align="right"><strong><%= Dis_str_money(fPSPricemTK)& DonviPrice%></strong></td>
                      <td align="right">&nbsp;</td>
                    </tr>
<%								
					else
					
					%>
                   <tr>
                     <td colspan="4" align="center">Chưa phát sinh</td>
                    </tr>
					<%end if%>					

                      <tr>
                      <td colspan="4" background="../../images/TabChinh.gif" height="30" class="CTieuDeNho" style="background-repeat:no-repeat">
					  &nbsp;Số điểm đã tích được					  </td>
                    </tr>
                      <tr>
                        <td align="center" style="<%=setStyleBorder(1,1,0,1)%>" height="28"><strong>Ngày đặt </strong></td>
                        <td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><strong>Đơn Hàng </strong></td>
                        <td align="center" style="<%=setStyleBorder(0,1,0,1)%>"><strong>Điểm</strong></td>
                        <td align="center" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
                      </tr>
					  <% 
					  	sql = "SELECT SanPhamUser.SanPhamUser_ID,SanPhamUser.CMND, SanPhamUser_Date, GetPoints FROM  SanPham_pay INNER JOIN SanPhamUser ON SanPham_pay.SanPhamUser_ID = SanPhamUser.SanPhamUser_ID WHERE (SanPhamUser_Status = 2) and (GetPoints > 0) and (SanPhamUser.CMND ='"& CMND &"')"
					 	Set rsPoint=Server.CreateObject("ADODB.Recordset")
						rsPoint.open sql,con,1
						tPoint = 0
					if not rsPoint.eof then
						do while not rsPoint.eof 
						DatePoint	= 	rsPoint("SanPhamUser_Date")
						iAddPoint	= 	rsPoint("GetPoints")
						tPoint		=	tPoint+iAddPoint
						tempSanPhamUser = "ea4bb1a25bfd1b838a8a940d02c8d8ec"&rsPoint("SanPhamUser_ID")&"d51dccca677d8049c356f8e5c830d7fc"
					  %>
                    <tr>
                     <td width="64" align="center" style="<%=setStyleBorder(1,1,0,1)%>" height="26"> <%=Day(DatePoint)%>/<%=Month(DatePoint)%>/<%=Year(DatePoint)%></td>
                      <td width="329" style="<%=setStyleBorder(0,1,0,1)%>">
					  <a href="ReportDH.asp?xboooooooo000000k=<%=tempSanPhamUser%>" class="CSubMenu" target="_blank">
			<%
				munb=1000+SanPhamUser_ID
				strTemp	="XB"+CStr(munb)
				Response.Write(strTemp)
			%>
					  </a></td>
                      <td width="108" align="center" style="<%=setStyleBorder(0,1,0,1)%>"><%=Dis_str_money(iAddPoint)%></td>
                      <td width="49" align="center" style="<%=setStyleBorder(0,1,0,1)%>">&nbsp;</td>
                    </tr>
					<%
						rsPoint.movenext
						loop
						set rsPoint = nothing
%>
                   <tr>
                      <td colspan="2" align="center"><strong>Tổng:</strong></td>
                      <td align="center"><strong><%= Dis_str_money(tPoint)%></strong> điểm</td>
                      <td align="center">&nbsp;</td>
                   </tr>
                   <tr>
                     <td colspan="2" align="right" style="<%=setStyleBorder(1,1,1,1)%>">
					 		Chuyển điểm thành tiền:					 </td>
					 <TD align="center" style="<%=setStyleBorder(0,1,1,1)%>">
					 	<a href="../ConvertPointToMoney.asp" target="_blank"><img src="../../images/icons/icon_coll_old_goods.gif" width="39" height="39" border="0" align="absmiddle" /></a> </TD>
					 <TD align="center" style="<%=setStyleBorder(0,1,1,1)%>">&nbsp;</TD>
                    </tr>
<%								
					else
					
					%>
                   <tr>
                     <td colspan="4" align="center">Chưa phát sinh</td>
                    </tr>
					<%end if%>
				<tr>
			  <td colspan="4" background="../../images/TabChinh.gif" height="30" class="CTieuDeNho" style="background-repeat:no-repeat">
			  &nbsp;Vấn tin tài khoản xbook</td>
			</tr>					
			<TR>
				<TD align="center" style="<%=setStyleBorder(1,1,0,1)%>">
				1</TD>
				<TD height="25" style="<%=setStyleBorder(0,1,0,1)%>">
				Tổng tiền tài khoản tăng </TD>
				<td style="<%=setStyleBorder(0,1,0,1)%>" align="right">+ <%=Dis_str_money(iTotalIni)%> </td>
			    <td style="<%=setStyleBorder(0,1,0,1)%>" align="right">&nbsp;</td>
			</TR>
			<TR>
			  <TD align="center" style="<%=setStyleBorder(1,1,0,1)%>">2.</TD>
			  <TD height="25" style="<%=setStyleBorder(0,1,0,1)%>">Số tiền thưởng tham Price đóng góp ý kiến </TD>
			  <td style="<%=setStyleBorder(0,1,0,1)%>" align="right">+ <%=Dis_str_money(iTotalIdea)%></td>
			  <td style="<%=setStyleBorder(0,1,0,1)%>" align="right">&nbsp;</td>
			</TR>
			<TR>
			  <TD align="center" style="<%=setStyleBorder(1,1,0,1)%>">3.</TD>
			  <TD height="25" style="<%=setStyleBorder(0,1,0,1)%>">Số tiền đã sử dụng </TD>
			  <td style="<%=setStyleBorder(0,1,0,1)%>" align="right">- <%=Dis_str_money(iTotalUser)%></td>
			  <td style="<%=setStyleBorder(0,1,0,1)%>" align="right">&nbsp;</td>
			</TR>
			<TR>
			  <TD height="25" colspan="2" align="center" style="<%=setStyleBorder(1,1,0,1)%>">Tổng tiền còn lại</TD>
			  <td style="<%=setStyleBorder(0,1,0,1)%>" align="right"><%= Dis_str_money(iTotalElse)%></td>
			  <td style="<%=setStyleBorder(0,1,0,1)%>" align="right">&nbsp;</td>
			</TR>
  </table>   
</body>
</html>
