<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="../css/styles.css" rel="stylesheet" />
<link href="../css/CommonSite.css" rel="stylesheet" />
<%
'''''''''''''''''''''''''''''''''''
Sub AddItemToCart(iItemID, iItemCount)
 	If dictCart.Exists(iItemID) Then
		dictCart(iItemID) = dictCart(iItemID) + 0
	Else
		dictCart.Add iItemID, iItemCount
	End If
End Sub
''''''''''''''''''''''''''''''''''''''''
Sub RemoveItemFromCart(iItemID, iItemCount)
 	If dictCart.Exists(iItemID) Then
		If dictCart(iItemID) <= iItemCount Then
			dictCart.Remove iItemID
		Else
			dictCart(iItemID) = dictCart(iItemID) - iItemCount
		End If
	End if
End Sub
''''''''''''''''''''''''''''''''''''''''
Sub ShowItemsInCart()
Dim Key
Dim aParameters 
Dim sTotal, sShipping
	%>
	<link href="../css/styles.css" rel="stylesheet" type="text/css" />
</head>
<form action="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=Thanhtoan&param=<%= NewsId %>&CatID=<%= CatID %> " name="fBanhang" method="post" >
	<div class="CTieuDeNho" style="width:100%;">Thông tin đặt hàng</div>	
    <TABLE width="100%" Border=0 align="center" CellPadding=0 CellSpacing=0  >
	    <TR  class="CTxtContent" >
			<TD  align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Tên sản phẩm </strong></TD>
			<TD  align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Loại bỏ</strong></TD>
			<TD  align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Đơn giá</strong></TD>
			<TD  align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>SL</strong></TD>
			<TD  align="center" style="<%=setStyleBorder(0,1,1,1)%>"><strong>Thành tiền</strong> </TD>
	    </TR>
	<%
	sTotal = 0
	sDem=0
	sTGiaBia	=	0
	For Each Key in dictCart
		aParameters = GetItemParameters(Key)
		%>
		<TR>
			<TD ALIGN="Left" class="CTxtContent" style="<%=setStyleBorder(0,1,0,1)%>">
			<%= aParameters(1) %>	
			</TD>
			<TD ALIGN="center" style="<%=setStyleBorder(0,1,0,1)%>">
		   <A HREF="./shopping.asp?action=del&param=<%= Key %>&CatID=<%= CatID %>&count=<%= dictCart(Key) %>" class="CSubMenu">Loại bỏ</A></TD>
			<TD ALIGN="Right" style="<%=setStyleBorder(0,1,0,1)%>">
			<input type="hidden" name="fDonGia<%=sDem%>" value="<%=aParameters(2)%>">
                <%sTGiaBia  = sTGiaBia	+  CLng(aParameters(4))*dictCart(Key) %>
 				<%= Dis_str_money(aParameters(2)) %>				</TD>
				
			<TD ALIGN="Right" style="<%=setStyleBorder(0,1,0,1)%>">
			<input name="Soluong<%=sDem%>"  value="<%=dictCart(Key) %>" size="2" maxlength="3" onBlur="javascrip:checkIsfBanhangNumber(this)" onKeyUp="javascrip: DisMoneyPriceHH();"  onKeyPress="ShowCapNhat();" style="width:25"  id="Soluong<%=sDem%>"></TD>
			<TD ALIGN="Right" style="<%=setStyleBorder(0,1,0,1)%>">
			<%
				tThanhTien = CLng(aParameters(2))*dictCart(Key)
			%>
			<input name="fThanhTien<%=sDem%>" type="text" class="CTextBox_NoBorder" size="7" readonly="readonly" style="text-align:right" value="<%=Dis_str_money(tThanhTien)%>">			</TD>
		</TR>
		<%
		sTotal = sTotal + (dictCart(Key) * CSng(aParameters(2)))
		SDem=SDem+1
	Next

	if sDem=0 then
	session("GioGang")=0
	%>
		<%Response.Write("<TR><TD COLSPAN=9 ALIGN=""center"" class=""CTxtContent""><B>Rỗng</B></TD></TR>")%>
	<%
		else
			payStatus =""
	%>
	
	<TR>
	  <TD COLSPAN=2 ALIGN="right" style="<%=setStyleBorder(1,1,0,1)%>"><b>Tổng tiền:</b></TD>
	  <TD COLSPAN=4 ALIGN="center" style="<%=setStyleBorder(0,1,0,1)%>">
	  <input type="hidden" name="NumSP" value="<%=SDem%>">
	  <input name="tThanhTien" type="text" class="CTextBox_NoBorder" size="8" readonly="readonly" value="<%=Dis_str_money(sTotal)%>"  >
	  	<%
		On Error Resume Next
		iCK =	100*(sTGiaBia - sTotal)/sTGiaBia
		iCK	=	round(iCK)
		Session("iCK")	= iCK	
		%>	  </TD>
	  </TR>

	<%
	session("GioGang")=sTotal
	End if%>
  </TABLE>

    
    <%Call  WriteFormUser()%>
     	
</form>	
	<%
End Sub
''''''''''''''''''''''''''''''''''''''''''''''


sub PhuongThucGiaoDich()
		fname 		    =	Request.Form("f_name")
		fmail 		    =	LCase(Request.Form("f_mail"))
		fmobile		    =	Request.Form("f_mobile")
        
        fwebsite		=	Request.Form("f_website")
        if InStr(fwebsite,"http://") > 0 then
            fwebsite = replace(fwebsite,"http://","")
            fwebsite = replace(fwebsite,"/","")
            fwebsite = "http://" + Trim(fwebsite)
        end if

        if InStr(fwebsite,"https://") > 0 then
            fwebsite = replace(fwebsite,"https://","")
            fwebsite = replace(fwebsite,"/","")
            fwebsite = "https://" + Trim(fwebsite)
        end if
         
        set rs=server.CreateObject("ADODB.Recordset")
        sql="select * from XSEOTitle where website like '%"& fwebsite &"%'"
        rs.Open sql, con, 1	
                  
        if not rs.eof then
	       response.Redirect("/shopping.asp?action=info&info=false")
        end if

        set rs=nothing 


        fcompany		=	Request.Form("f_company")
		faddress 	    =	Trim(Request.Form("f_address"))
        faddress_tax    =	Trim(Request.Form("f_address_tax"))
        ftel 	        =	Trim(Request.Form("f_tel"))				
		fmst		    =	Request.Form("f_mst")
        fcodeoff 	    =	Trim(Request.Form("f_codeoff"))				
		
        IDPrice		=	GetNumeric(Request.Form("RPrice"),0)

        saleoff = false
        AgentID =   0
        if fcodeoff <> "" then
            sql = "SELECT  * FROM NhanVien where salecode ='" & fcodeoff &"'"
            Set rs = Server.CreateObject("ADODB.Recordset")
            rs.open sql,Con,1
            If not rs.eof Then
                saleoff    =    true
                AgentID    =    rs("NhanVienID")
            end if
        end if

        Price   = 0
        NetPrice = 0
        set rs=server.CreateObject("ADODB.Recordset")
        sql="SELECT * FROM XSEOPrice where ID = '"& IDPrice &"'"
        rs.open sql,con,1 
        If not rs.eof Then
            NetPrice    =  rs("Price") 
            if saleoff = true then
                Price   =   rs("PriceOff")
            else
                Price   =   NetPrice
            end if            
            fMonth      =   rs("Month")
            fDownload   =   rs("Download")
        end if

        StartDate   =   now
        EndDate     =   now + 30*fMonth

        set rs=server.CreateObject("ADODB.Recordset")
        sql="INSERT INTO XSEOTitle(website, activation, TryVersion, StartDate, EndDate) values "
        sql= sql +	"('" & fwebsite & "','true','true','"& StartDate &"','"& EndDate &"')"
        rs.Open sql, con, 1	
        set rs=nothing

        IDWeb   =   GetMaxID("XSEOTitle","ID","")   -   1
        
        set rs=server.CreateObject("ADODB.Recordset")
        sql="INSERT INTO SanPhamUser(Name, IDWeb, Mobile, Email, Tel, Address, Company, Masothue, AddressTax, Status,Codeoff,GhiChu) values "
        sql= sql +	"(N'" & fname & "','"& IDWeb &"',N'"& fmobile & "',N'"& fmail &"','"& ftel &"',N'"& faddress &"',N'"& fcompany &"'"
        sql= sql +	",'"& fmst & "',N'"& faddress_tax &"','1','"  & fcodeoff & "',N'Khách đang đặt dở')"

        rs.Open sql, con, 1	
        set rs=nothing

        SanPhamUser_ID  =   GetMaxId("SanPhamUser","SanPhamUser_ID","")-1

        
        set rs=server.CreateObject("ADODB.Recordset")
        sql="INSERT INTO SanPham_Pay(SanPhamUser_ID, Price, DeductAccount, AgentID, RateAgent, Note) values "
        sql= sql +	"('" & SanPhamUser_ID & "','"& Price &"','"& DeductAccount &"','"& AgentID &"','"& RateAgent &"',N'"& Note &"')"
        rs.Open sql, con, 1	
        set rs=nothing
	
%>
<form action="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=checkout" name="fBanhang" method="post" >
   
	<input type="hidden" value="<%=fname%>" name="fname" />
    <input type="hidden" value="<%=fmail%>" name="fmail" />
    <input type="hidden" value="<%=fmobile%>" name="fmobile" />
    <input type="hidden" value="<%=fwebsite%>" name="fwebsite" />
    <input type="hidden" value="<%=fcompany%>" name="fcompany" />
    <input type="hidden" value="<%=faddress%>" name="faddress" />
    <input type="hidden" value="<%=faddress_tax%>" name="faddress_tax" />
    <input type="hidden" value="<%=ftel%>" name="ftel" />
    <input type="hidden" value="<%=fmst%>" name="fmst" />
    <input type="hidden" value="<%=fDownload%>" name="fDownload" />
    
    
    <input name="SanPhamUser_ID" type="hidden" value="<%=SanPhamUser_ID%>" />
	<input name="IDWeb" type="hidden" value="<%=IDWeb%>" />
	
    <input type="hidden" value="<%=AgentID%>" name="AgentID" />
	<input type="hidden" value="<%=Price%>" name="Price">

  <table border="0"  class="box-border-no-fix" style="width:800px; padding:2px; background-color:#ffffff;margin:auto;">
	
    <tr>
      <td class="CTieuDe" height="35px" > 
	 Thanh toán:</td>
    </tr>
      <tr>
          <td class="CTxtContent">
              <%if Price <> 0 then
                    call TKBank()
                else
                  Response.Write(fOnlyCallName(fname)&" đang dùng phiên bản dùng thử, tuy nhiên để có được sự trải nhiệm tốt nhất "& fXungHo(fName)&" có thể chuyển thành phiên bản đầy đủ bất cứ khi nào nếu mã website của "& fXungHo(fName)&" được ghi nhận đã thanh toán. <br> Điều này có thể được thiết lập về sau")
                end if
                %>
          </td>
      </tr>
      <tr>
          <td class="CTxtContent" style="padding:10px;">
              <%if Price <> 0 then%>
              <input name="ichoicePayBank" type="radio" value="-1" onclick="javascript:ShowInfoBank('<%=strInfoBank%>');"/> Tôi muốn thanh toán tiền mặt với người của XSEO liên hệ thu trực tiếp.<br />
              <input name="ichoicePayBank" type="radio" value="-2" onclick="javascript:ShowInfoBank('<%=strInfoBank%>');"/> Tôi sẽ đến văn phòng công ty nộp tiền mặt trực tiếp.
              <%end if %>
              <div style ="width:30%;float:right;border-color:#ff6a00;" class="border-box-img">
                 <div class="CTieuDeNho" align="center" style="color:#514ffa;"> HÓA ĐƠN</div>
                  <div>Tổng thành tiền</div>
                   <div class="CTieuDe" align="center" style="color:#ff6a00;"> <%=Dis_str_money(Price) %> đ</div>
              </div>
                                <%
    If saleoff = true Then
        %>    <div style ="width:50%;float:left;border-color:#ff6a00;" class="box-border-no-fix">
                  Chúc mừng <%=fOnlyCallName(fname)%> mã giảm giá đã được chấp thuận! <br>
              Hóa đơn được giảm
               <div class="CTieuDe" align="center" style="color:#ff6a00;">Giảm <%=Dis_str_money(NetPrice-Price) %> đ</div>
            </div>
        <%end if%> 
          </td>
      </tr>
   
    <tr>
      <td align="center" >
          <br /><br />
                <%if Price <> 0 then%>
                   <input name="Button" type="button" onClick="javascript: if (document.fBanhang.ichoicePayBank.value != '') { submit();} else alert('Xin hãy chọn một cách thức để thanh toán');  " value="Tiếp tục" style="background-color:#ff6a00; width:128px;height:128px; color:white; font-size:18px;letter-spacing:2px;text-overflow:clip;" class="ButtonCircle" /> 
                <%else%>
                  <input name="Button" type="button" onClick="javascript: submit();" value="Tiếp tục" style="background-color:#ff6a00; width:128px;height:128px; color:white; font-size:18px;letter-spacing:2px;text-overflow:clip;" class="ButtonCircle" />
                <%end if%>
           </td>
    </tr>

  </table>
</form>
  <%
end sub
%>
	
<%sub TKBank()%>
<table  border="0" align="center" cellpadding="2" cellspacing="2" class="CTxtContent" style="border:#CCCCCC solid 1;" width="750">
  <tr>
    <td colspan="10" class="CTieuDeNhoNho" style="border-bottom:#CCCCCC solid 1px;">
	  Mời quý khách chọn chuyển tiền vào một trong các Ngân hàng sau: </td>
  </tr>
  <tr>
    <td width="70" align="center"><img src="../images/Logo/ACB.jpg" width="30" height="30" />        </td>
    <td width="70" align="center"><img src="../images/Logo/AGRIBANK.jpg" width="30" height="30" />   </td>
   	<td width="70" align="center"><img src="../images/Logo/BIDV.jpg" width="30" height="30" /></td>
    <td width="70" align="center"><img src="../images/Logo/Techcombank.jpg" width="70" height="30" /></td>
    <td width="70" align="center"><img src="../images/Logo/Vietcombank.jpg" width="70" height="30" /></td>
   	<td width="70" align="center"><img src="../images/Logo/VIETINBANK.jpg" width="70" height="30" /> </td>
  </tr>
  <tr>
    <td width="70" align="center">
	<%
		strInfoBank	=	"Tại Ngân hàng Á Châu (ACB) CN Trần Đại Nghĩa(HN)- Số TK: 70539919  - Chủ TK: Nguyễn Văn Tuân"
	%>
	<input name="ichoicePayBank" type="radio" value="1"  onclick="javascript:ShowInfoBank('<%=strInfoBank%>');"/>	</td>
    <td width="70" align="center">
	<%
		strInfoBank	=	"Tại NH Nông nghiệp & PT Nông thôn VN (AGRIBANK) CN Lê Thanh Nghị(HN) - Số TK: 1303206052942  - Chủ TK: Nguyễn Văn Tuân"
	%>
	<input name="ichoicePayBank" type="radio" value="2" onclick="javascript:ShowInfoBank('<%=strInfoBank%>');"/>	</td>
   	<td width="70" align="center">
	<%
		strInfoBank	=	"Tại NH Đầu tư & PT Việt Nam (BIDV) CN Hà Thành - Số TK: 12210000124971  - Chủ TK: Nguyễn Văn Tuân"
	%>	
	<input name="ichoicePayBank" type="radio" value="3" onclick="javascript:ShowInfoBank('<%=strInfoBank%>');"/>	</td>
    <td width="70" align="center"><%
		strInfoBank	=	"Tại NH TMCP Kỹ thương Việt Nam (TECHCOMBANK) CN Hoàn Kiếm(HN) - Số TK: 10820553655019  - Chủ TK: Nguyễn Văn Tuân"
	%>
      <input name="ichoicePayBank" type="radio" value="6" onclick="javascript:ShowInfoBank('<%=strInfoBank%>');"/></td>
    <td width="70" align="center"><%
		strInfoBank	=	"Tại NH TMCP Ngoại thương Việt Nam (VIETCOMBANK) CN Hà nội - Số TK: 0021002013541 - Chủ TK: Nguyễn Văn Tuân"
	%>
      <input name="ichoicePayBank" type="radio" value="7" onclick="javascript:ShowInfoBank('<%=strInfoBank%>');"/></td>
   	<td width="70" align="center"><%
		strInfoBank	=	"Tại NH TMCP Công thương VN(VIETINBANK) CN Hoàng Mai(HN) - Số TK: 711A21957483  - Chủ TK: Nguyễn Văn Tuân"
	%>
      <input name="ichoicePayBank" type="radio" value="8" onclick="javascript:ShowInfoBank('<%=strInfoBank%>');"/></td>
  </tr>
     </table>
    <div class="CTxtContent" align="center" > 
        <input name="txtChoieBank" type="text" class="CTextBox_NoBorder" id="txtChoieBank" size="120" readonly="true" style="text-align:center;">
    </div>
        <div class="CTxtContent" align="right">
            <input name="bHelpBank" type="button" value="  ?  " onclick="javascript:ShowHelpBank();" style="height:30px;"/><input type="hidden" name="hShowHelp" value="0" />
        </div> 
    <div id="ShowHelpBank" style="display:none; background-color:#fcffad;padding:5px; margin:auto;" class="border-box-text" >
        <span class="CTxtContent">
           Mời quý khách chọn một trong những ngân hàng trên để giao dịch, sao cho thuận tiện cho việc chuyển khoản của quý khách. Sau đó quý khách lưu thông tin số tài khoản và chủ tài khoản của chúng tôi.<br />
	       Để thực hiện việc chuyển tiền vào tài khoản quý khách có thể chuyển tiền bằng các hình thức sau:<br />
            1. Sau khi việc đặt hàng tại XSEO kết thúc quý khách có thể chuyển tiền trực tuyến từ hệ thống website của các ngân hàng vào tài khoản của chúng tôi<br />
	       2. Quý khách tới địa điểm giao dịch gần nhất của Ngân hàng  đã chọn sau đó viết phiếu chuyển tiền.<br />
	       3. Hoặc quý khách tới máy ATM gần nhất của hệ thống Ngân hàng  đã chọn, sau đó thực hiện giao dịch chuyển khoản (trường hợp khách có thẻ ATM).<br />
	       3. Hoặc quý khách thực hiện các giao dịch khác theo tiện ích của từng đơn vị như: Mobile banking,  
        OTP....
        <br />
        Nội dung ghi thanh toán &quot;Thanh toan xseo mã ....&quot; <b>mã</b> website sẽ được nhận tại bước tiếp theo<br />
	       <b>Lưu ý</b>: Trước khi thực hiện chuyển tiền quý khách vẫn có thể cài đặt XSEO nhưng mặc định là phiên bản chưa được kích hoạt, ngay sau khi kết thúc giao dịch chuyển tiền của quý khách chúng tôi sẽ nhận được thông báo của hệ thống Ngân hàng. thì hệ thống phần mềm sẽ tự động kích hoạt trạng thái cho mã website của bạn mà không cần phải liên lạc với chúng tôi.<br />
            Đối với trường hợp cần viết Hóa đơn, chúng tôi sẽ chủ động chuyển phát nhanh tới địa chỉ quý khách yêu cầu.
	    <div  align="right"> Xin trân thành cảm ơn quý khách.quý khách.</div>  

        </span>
	
    </div>


<script>
        function ShowHelpBank() {

        if (document.fBanhang.hShowHelp.value == 0) {
            document.getElementById("ShowHelpBank").style.display = "";
            document.fBanhang.hShowHelp.value = 1;
            document.fBanhang.bHelpBank.value = ' Ẩn Hướng Dẫn ';
        }
        else {
            document.getElementById("ShowHelpBank").style.display = "none";
            document.fBanhang.hShowHelp.value = 0;
            document.fBanhang.bHelpBank.value = ' ? ';
        }
    }
        function ShowInfoBank(info) {
        document.fBanhang.txtChoieBank.value = String(info);
    }
</script>
<%end sub%>	
	

<%sub infocus()
    if Session("email")<> "" then
        sql="SELECT * FROM Account WHERE (Email = '" & Session("email") & "')"
        Set rs = Server.CreateObject("ADODB.Recordset")
    	rs.open sql,con,1
	    if not rs.eof then
           fname    =   rs("Name")                 
           fmobile  =   rs("Mobile")
           fEmail   =   rs("Email")
           fAddress =   rs("Address")
        end if
    end if
 %>
<table style="background-color:#ffffff; margin:auto;" class="box-border-no-fix">
    <tr>
        <td>
       <div style="width:100%; margin:auto auto;">
            <% Call Ngang_Menu(0,"VN",0)%>
       </div>
        </td>
    </tr>
            <tr>
              <td>
                                   
		<form name="form_reg" method="post" action="shopping.asp?action=payment">
		<table  align="center"  class="CTxtContent" style="padding:2px;width:800px; ">
			<tr>
			  <td  class="CTieuDe" align="center" colspan="2" >
                  <img src="images/icons/icon_customer1.gif" width="48" height="48" align="absmiddle">ĐĂNG KÝ
               </td>
			</tr>
			<tr>
			  <td align="center"  class="CTxtContent" colspan="2">
                  <table style="width:100%;">
                      <tr>
                          <%
                            set rs1=server.CreateObject("ADODB.Recordset")
                            sql="SELECT  TOP (100) PERCENT ID, Note,title, Price, Month, Description FROM XSEOPrice ORDER BY Month"
                            rs1.open sql,con,1 
                            stt = 0
                            total_news = rs1.recordcount-1
                            Redim arTitle(total_news,2)
                            Do while not rs1.eof
                              title =   rs1("title")
                              arTitle(stt,1)	=	rs1("Note")
                              arTitle(stt,2)	=	rs1("Description")
                              %>
                          <td style="padding:5px; margin:auto;" align="center" class="box-border-no-fix">
                              <%=title%>
                              <br />
                              <input id="RPrice<%=stt%>" type="radio" name="RPrice" value="<%=rs1("ID")%>" onClick="javascript:AD_OnMouseOver(<%=stt%>,<%=total_news%>,'list_document','RPrice')" />
                          </td>
                            <%	
                                stt=stt + 1
                                rs1.MoveNext
                            Loop
                            rs1.close %>
                      </tr>
                  </table>
                </td>
			</tr>
			<tr>
                <td colspan="2" align="center" class="CTxtContent">
                    <%for j=0 to total_news%>	
                    <table style="display:none; width:450px;background-color:#2957A4;color:#ffffff; " id="list_document<%=j%>"  class="border-box-text">
                        <tr>
                            <td class="CTieuDe" style="color:#ffffff;" align="center"><%=arTitle(j,1)%></td>
                        </tr>
                        <tr>
                            <td class="CTxtContent" style="color:#ffffff;"><%=arTitle(j,2)%></td>
                        </tr>
                    </table>
                    <%next%>
                </td>

			</tr>
			<tr>
			  <td align="right"  class="CTxtContent">&nbsp;</td>
			  <td class="CTxtContent">&nbsp;</td>
			</tr>
            <%if Request.QueryString("info") = "false" then %>
            <tr>
                <td class="CTieuDeNhoNho" style="color:#ff0000;" colspan="2">Thông báo lỗi: website đã được đăng ký, hãy đăng ký website khác hoặc liên hệ với chúng tôi để được biết thêm chi tiết</td>
            </tr>
			<%end if %>
			<tr>
			  <td align="right"  class="CTxtContent">Họ và tên</td>
			  <td class="CTxtContent"><input name="f_name" type="text" id="f_name" value="<%=fname%>" size="45" class="CTextBoxUnder"></td>
			</tr>
			
			<tr>
			  <td align="right" class="CTxtContent">Di động:</td>
			  <td class="CTxtContent">
			      <input name="f_mobile" type="text" id="f_mobile" size="45" class="CTextBoxUnder" value="<%=fmobile%>"></td>
			  
			</tr>
			
			<tr>
			  <td align="right" class="CTxtContent">Email:</td>
			  <td class="CTxtContent"><input name="f_mail" type="text" id="f_mail" value="<%=fEmail%>" size="45" class="CTextBoxUnder"></td>
			 
			</tr>
            <tr>
			  <td align="right" class="CTxtContent">Website kích hoạt:</td>
			  <td class="CTxtContent">
			  <input name="f_website" id="f_website" size="45" class="CTextBoxUnder" value="<%=fwebsite%>"></td>
			 
			</tr>
            <tr>
			  <td align="right" class="CTxtContent">Mã giảm giá: </td>
			  <td class="CTxtContent">
			  <input name="f_codeoff" id="f_codeoff" size="45" class="CTextBoxUnder" value="<%=f_codeoff%>"></td>
			 
			</tr>
            <tr>
                <td></td>
                <td >
                    <input id="CheckExp" type="checkbox" name="CheckExp" value="1" onClick="javascript: showother();"  /> Thêm thông tin mở rộng 
                </td>
            </tr>
            </table>
			<div style="display:none;" id="list_exp">
                <table align="center" style="padding:2px;width:800px; " class="CTxtContent">  
			        <tr>
			          <td align="right">Tên công ty:</td>
			          <td class="CTxtContent">
			          <input name="f_company" id="f_company" size="45" class="CTextBoxUnder" value="<%=fCompany%>"></td>
			        </tr>
			        <tr>
			          <td align="right">Mã số thuế:</td>
			          <td class="CTxtContent">
			          <input name="f_mst" id="f_mst" size="45" class="CTextBoxUnder" value="<%=fmst%>"></td>
			        </tr>
			
    		        <tr>
			          <td align="right">Địa chỉ thuế: </td>
			          <td class="CTxtContent">
			          <input name="f_address" id="f_address" size="45" class="CTextBoxUnder" value="<%=fAddress%>"></td>
    		        </tr>
    		        <tr>
			          <td align="right">Địa chỉ nhận hóa đơn: </td>
			          <td class="CTxtContent">
			          <input name="f_address_tax" id="f_address_tax" size="45" class="CTextBoxUnder" value="<%=fAddress_tax%>"></td>
    		        </tr>
                    <tr>
			          <td align="right">Điện thoại công ty:</td>
			          <td class="CTxtContent"><input name="f_tel" type="text" id="f_tel" size="45" class="CTextBoxUnder" value="<%=fTel%>"></td>
			        </tr>
		        </table> 
                </div> 
            <center>
                <input name="Button" type="button" id="Submit" onClick="Checkreg()" value="Đăng ký" style="background-color:#ff6a00; width:128px;height:128px; color:white; font-size:18px;letter-spacing:2px;text-overflow:clip;" class="ButtonCircle" />  
            </center>
            
		  </form>
				  
              </td>
            </tr>
</table>
<%end sub %>

<%
Sub PlaceOrder()
		fname 		    =	Request.Form("fname")
		fmail 		    =	Request.Form("fmail")
		fmobile		    =	Request.Form("fmobile")
        fwebsite		=	Request.Form("fwebsite")
        
        ExistEmail      =   false
        sql="SELECT * FROM Account WHERE (Email = '" & fmail & "')"
        Set rs = Server.CreateObject("ADODB.Recordset")
    	rs.open sql,con,1
	    if not rs.eof then
            ExistEmail  =   true
        end if
    
        fcompany		=	Request.Form("fcompany")
		faddress 	    =	Trim(Request.Form("faddress"))
        faddress_tax    =	Trim(Request.Form("faddress_tax"))
        ftel 	        =	Trim(Request.Form("ftel"))				
		fmst		    =	Request.Form("fmst")
        Price 	        =	Request.Form("Price")
        AgentID 	    =	Request.Form("AgentID")
        IDWeb 	        =	Request.Form("IDWeb")
        fDownload       =   Request.Form("fDownload")
        SanPhamUser_ID  =   Request.Form("SanPhamUser_ID")
%>
<form action="<%=Request.ServerVariables("SCRIPT_NAME")%>?action=End" name="fBanhang" method="post" onSubmit="checkme()">
    <input type="hidden" value="<%=fname%>" name="fname" />
    <input name="SanPhamUser_ID" type="hidden" value="<%=SanPhamUser_ID%>" />
    <input name="fDownload" type="hidden" value="<%=fDownload%>" />
    <input name="IDWeb" type="hidden" value="<%=IDWeb%>" />
    <input name="fmail" type="hidden" value="<%=fmail%>" />
    <input name="fmobile" type="hidden" value="<%=fmobile%>" />
    <input name="faddress" type="hidden" value="<%=faddress%>" />

    <table border="0"  class="box-border-no-fix" style="width:800px; padding:2px; background-color:#ffffff;margin:auto;">
    <tr>
        <td align="center">
            <img src="<%=Logo %>" alt="xseo.vn"  height="64" />
        </td>
    </tr>
    
    <%if Session("Email")="" and ExistEmail = false and (Price=0 or Price>2000000) then %>
    <tr >
        <td class="CTxtContent">
       <b><%=fOnlyCallName(fname)%> thân mến!</b> <br />
        <%=fXungHo(fname) %> đã đăng ký phiên bản XSEO cao cấp vì vậy sẽ được cấp tài khoản quản lý từ xa, mà chỉ có <%=fXungHo(fname) %> mới được quyền thiết lập XSEO chạy trên các máy tính khác nhau theo cách của mình mong muốn.<br />
            Để thuận tiện phần mềm XSEO sẽ lấy thông tin email đăng nhập là:<b><%=fmail %></b> và cần khởi tạo password ban đầu tại đây: <br />
            <input id="Password1" type="text"  size="35"  name="Password" class ="CTextBoxUnder"/>
            <input id="Checkbox1" name="Checkbox1" type="checkbox" onclick="javascript: if (document.fBanhang.Checkbox1.checked == true) {document.fBanhang.Password1.type    =   'password';} else {document.fBanhang.Password1.type    =   'text';}" /> Hiện/ẩn Password
             <hr />
        </td>

    </tr>
    <%end if %>
   <%if ExistEmail = true and Session("Email")="" and (Price=0 or Price>2000000) then %>
    <tr >
        <td class="CTxtContent">
       <b><%=fOnlyCallName(fname)%> thân mến!</b> <br />
        Với email này <%=fmail%>&nbsp;<%=fXungHo(fname) %> đã có tài khoản tại XSEO.vn vì vậy sau khi đăng ký thành công <%=fXungHo(fname) %> hãy đăng nhập để thiết lập chạy XSEO
        </td>

    </tr>
    <%end if %>
    <tr>
      <td > 
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
    <tr>
      <td colspan="2" align="center"  class="CTieuDe">THÔNG TIN ĐĂNG KÝ</td>
    </tr>
    <tr>
      <td width="13%" align="right" >Họ và Tên :</td>
      <td width="35%" ><b><%=fname%></b> </td>
    </tr>
    <tr>
      <td align="right" >Email :</td>
      <td ><strong><%=fmail%></strong></td>
    </tr>
    <tr>
      <td align="right" >Mobile:</td>
      <td ><strong><%=fmobile%></strong></td>
    </tr>
      <%if fmst <>"" then %>
      <tr>
          <td colspan="2" class="box-border-no-fix">
              <b>Xác nhận lại thông tin viết hóa đơn<br /></b>
             <i>Tên công ty:</i> <%=fcompany%><br />
              <i>Mã số thuế:</i> <%=fmst%><br />
              <i>Địa chỉ thuế:</i><%=faddress_tax%><br />
              <i>Địa chỉ nhân hóa đơn:</i> <%=faddress%><br />
              <i>Điện thoại:</i> <%=ftel %>
          </td>
      </tr>
      <%end if %>
    <tr>
      <td align="right" >Website kích hoạt:</td>
      <td ><strong><%=fwebsite%></strong></td>
    </tr>
      <tr>
          <td align="right" >Mã kích hoạt:</td>
          <td class="CTieuDeNho"><b><%=(1000+IDWeb)%></b>
          </td>
      </tr>
      <tr><td class="CTxtContent" colspan="2">
         Mã <b><%=(1000+IDWeb)%></b> đây là mã để kích hoạt website của <%=fXungHo(fname)%>, vì vậy <%=fOnlyCallName(fname)%>  cần ghi nhớ để điền thông tin trong quá trình cài đặt phần mềm
                <br /> Bộ cài được download  <a href="<%=fDownload%>" class="CTieuDeNhoNho"> TẠI ĐÂY </a>
          </td></tr>
  </table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="CTxtContent">
<%if isKhuyenMai(now)<>0 and Session("ChoiceKM")<>0 then%>
	<TR>
	  <TD  ALIGN="left" class="CTxtContent">
	  <strong>Khuyến mãi: </strong><br>
		<%call showSubKhuyenMai(Session("ChoiceKM"))%>
		<br>	  </TD>
    </TR>
<%end if%>	
	<%
		ichoicePayBank 	= 	Clng(Request.Form("ichoicePayBank"))
		if ichoicePayBank > 0 then
            strTToanVanChuyen="<b>Hình thức thanh toán</b>:" + fOnlyCallName(fname) + " vui lòng chuyển tiền vào tài khoản ngân hàng:<br>"&Trim(Request.Form("txtChoieBank"))&".<br>Tổng tiền thanh toán là: <b>"& Dis_str_money(Price)  &" đ</b><Br>Nội dung ghi: Thanh toán XSEO mã "& (1000+IDWeb)
        elseif ichoicePayBank = - 1 or ichoicePayBank = - 2 then
	       strThoiGian = "<b>Người XSEO tư vấn: </b>"
            sql = "Select top 1 * from ThongBao where KichHoat <> 0 and (Subjects = 1 or Subjects=2) and Type = 1 and (DATEDIFF(""n"",'"&now&"',FromDate)<= 0) AND (DATEDIFF(""n"",'"&now&"',ToDate) >= 0)   ORDER BY FromDate DESC "
            set rsTB = Server.CreateObject("ADODB.recordset")
            rsTB.Open sql,con,1
            if not rsTB.eof then
	            FromDate = rsTB("FromDate")
	            ToDate	= rsTB("ToDate")
            end iF			
            if isDate(FromDate) = true and isDate(ToDate)=true and DATEDIFF("n",FromDate,now)>= 0 and DATEDIFF("n",ToDate,now)<= 0 then
                strThoiGian =strThoiGian + "XSEO sẽ liên hệ lại với "& fOnlyCallName(fname) &" sau "&hour(ToDate)&"h:"&minute(ToDate)&"p&nbsp;ngày "&Day(ToDate)&"/"&Month(ToDate)&"/"&Year(ToDate)	
            else
                strThoiGian =strThoiGian + "XSEO sẽ liên lại lại với "& fOnlyCallName(fname) &" trong vòng 24h trừ chủ nhật để tư vấn và giúp hoàn thiện thủ tục"
            end iF
        end if
	%>
    <%if Len(strTToanVanChuyen&strThoiGian) > 10 then%>
	<TR>
	  <TD  ALIGN="left" class="CTxtContent">
          
          <div class="box-border-no-fix" style="background-color:#5a5a5a; color:#ffffff; padding:5px;">
            <%=strTToanVanChuyen+strThoiGian%>
          </div> 	    
          XSEO chỉ kích hoạt đầy đủ tính năng sau khi nhận được thông báo chuyển tiền và việc kích hoạt được tự động mà không cần <%=fXungHo(fname)%> tác động trên hệ thống, kính mong <%=fXungHo(fname)%> vui lòng sớm chuyển tiền vào tài khoản XSEO để được ưu đãi tốt nhất.
	</TD>
</TR>
    <%end if %>

	<TR>
	  <TD  ALIGN="left" class="CTxtContent" valign="top">
Yêu cầu khác (nếu có):<br />
              <textarea name="IdeaOther" cols="50" rows="3" id="textarea" onkeyup="initTyper(this)" class="border-box-text"> </textarea>
      </TD>
    </TR>	
	<tr>
	  	<td align="center">
			  <input type="submit" value="Đồng ý"  name="Dongy" style="background-color:#ff6a00; width:128px;height:128px; color:white; font-size:18px;letter-spacing:2px;text-overflow:clip;" class="ButtonCircle"/>		</td>
    </tr>
</table>
          </td>
        </tr>
          </table>
</form>		
	<%
End Sub

%>
