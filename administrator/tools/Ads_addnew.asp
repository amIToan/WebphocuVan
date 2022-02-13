<%@  language="VBSCRIPT" codepage="65001" %>
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
	
	if Request.Querystring("action")="Insert" then
		sError=False
		Set Upload = Server.CreateObject("Persits.Upload")

		Upload.SetMaxSize 1000000, True 'Dat kich co upload la` 1MB
		Upload.codepage=65001
		Upload.Save
		'Ads_id
		Ads_id=GetMaxId("Ads", "Ads_id", "")
		'Ads_Title
		Ads_Title=ReplaceHTMLToText(Upload.form("Ads_Title"))
		if Ads_Title="" then
			sError=True
		end if
		'Ads_Link
		Ads_Link=ReplaceHTMLToText(Upload.form("Ads_Link"))
		
		'Ads_Images
		set Ads_Images = Upload.Files("Ads_ImagesPath")
		If Ads_Images Is Nothing Then
			Ads_ImagesPath=""
		else
		   Filetype = Right(Ads_Images.Filename,len(Ads_Images.Filename)-Instr(Ads_Images.Filename,"."))
		   if Lcase(Filetype)<>"jpg" and Lcase(Filetype)<>"gif" and Lcase(Filetype)<>"jpeg" and Lcase(Filetype)<>"swf" and Lcase(Filetype)<>"png"then
				sError=True
				sAsd_ImagesPath="Không là file ảnh"
		   else
		   		Ads_ImagesPath="Ads_ImagesPath_" & Ads_id & "." & Filetype
		   end if
		End If



    


    idcolor1=Trim(Upload.form("idcolor_tex1"))
    idcolor2=Trim(Upload.form("idcolor_tex2"))


		'Width x Height
		if Ads_ImagesPath="" then
		'Nếu kô có Banner hoặc Icon, Lấy chiều cao hiển thị của quảng cáo text
			Ads_height=GetNumeric(Upload.form("Ads_height"),0)
			Ads_width=0
			Ads_Type=2
		elseif Lcase(Filetype)="swf" then
		'Lấy chiều rộng và chiều cao của flash
			Ads_height=GetNumeric(Upload.form("Ads_height"),0)
			Ads_width=GetNumeric(Upload.form("Ads_width"),0)
			if Ads_height=0 or Ads_width=0 then
				sError=True
				sAds_width="flash: Bắt buộc"
			end if
			Ads_Type=1
		else
		'Lấy chiều rộng và chiều cao của ảnh
			Ads_height=Ads_Images.ImageHeight 
			Ads_width=Ads_Images.ImageWidth
			Ads_Type=0
		end if
		'Ads_Position
		Ads_Position=GetNumeric(Upload.form("Ads_Position"),-1)
		if Clng(Ads_Position)=-1 then
			sError=True
		end if
		
		Ads_url=Clng(Upload.form("Ads_url"))
		'StatusId
		StatusId=ReplaceHTMLToText(Upload.form("StatusId"))
		if statusId<>"eded" and statusId<>"apap" then
			sError=True
		end if
		'CategoryId
		CategoryId=GetNumeric(Upload.form("CategoryId"),-2)
		if CategoryId=-2 then
			sError=True
		end if
		'Ads_OnlineChildren
		Ads_OnlineChildren=GetNumeric(Upload.form("Ads_OnlineChildren"),0)
		'Ads_Note
		Ads_Note=ReplaceHTMLToText(Upload.form("Ads_Note"))
		
		if not sError then
			Dim rs
			set rs=server.CreateObject("ADODB.Recordset")
			if Ads_ImagesPath <>"" then 
				Ads_Images.SaveAs Path & "\" & Ads_ImagesPath
			end if
			sql="INSERT INTO Ads (Ads_id,Ads_Title,Ads_Link,Ads_ImagesPath,Ads_Type,Ads_width," &_
				"Ads_height,Ads_Position,StatusId,Ads_Creator,Ads_Note,idcolor_tex1,idcolor_tex2, Ads_url) values " &_
				"(" & Ads_id &_
				",N'" & Ads_Title & "'" &_
				",'" & Ads_Link & "'" &_
				",'" & Ads_ImagesPath & "'" &_
				"," & Ads_Type &_
				"," & Ads_width &_
				"," & Ads_height &_
				"," & Ads_Position &_
				",'" & StatusId & "'" &_
				",'" & Session("user") & "'" &_
				",N'" & Ads_Note & "'" &_
				",N'" & idcolor1 & "'" &_
				",N'" & idcolor2 & "'" &_
				",'" & Ads_url & "')"
    
				
			response.write sql & "<br>"
			'response.end
			rs.open sql,con,1
			if CategoryId=-1 then
			'Insert vào nhiều chuyên mục khác nhau
				set rs=nothing
				set upload=nothing
				response.redirect ("Ads_MultiCat.asp?id=" & Ads_id)
			
			else
				sql="INSERT INTO AdsDistribution (Ads_id,CategoryId, Ads_OnlineChildren, Ads_Order) values " &_
				"(" & Ads_id &_
				"," & CategoryId &_
				"," & Ads_OnlineChildren &_
				"," & GetMaxId("AdsDistribution", "Ads_Order", "") & ")"
			end if

			rs.open sql,con,1
			'rs.open sql,con,1 : Quyen Ghi
			'rs.open sql,con,3 : Quyen Doc
			response.Write "<script language=""JavaScript"">" & vbNewline &_
			"<!--" & vbNewline &_
				"window.opener.location.reload();" & vbNewline &_
				"window.close();" & vbNewline &_
			"//-->" & vbNewline &_
			"</script>" & vbNewline
			set rs=nothing
			set upload=nothing
			response.End()
		end if
	else
		CategoryId=-2
		Ads_Position=-1
		StatusId=-1
	end if 'Of if Request.Form("action")="Insert" then
%>
<html>
<head>
    <title><%=PAGE_TITLE%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body>
    <form action="<%=request.ServerVariables("SCRIPT_NAME")%>?action=Insert" method="post" enctype="multipart/form-data" name="fInsertEvent">
        <table border="0" align="center" cellpadding="2" cellspacing="2" class=" Tb-input Tb-in ">
            <tr>
                <th colspan="3">Tạo mới Banner &amp; Icon quảng cáo</th>
            </tr>
            <tr>
                <td style="width: 24%;"> Tiêu đề: </td>
                <td style="width: 76%;">

                    <table class="tb-it">
                        <tr>
                                                           
                           
							<td>
                                <input name="Ads_Title" type="text" id="" size="35" maxlength="150" value="<%=Ads_Title%>"></td>
                        </tr>
                    </table>




                </td>
                <td style="color: #F00;">*</td>
            </tr>
            <tr>
                <td>Đường Link:</td>
                <td>
                    <table class="tb-it">
                        <tr>
                            <td>
                                <input name="Ads_Link" type="text" id="Ads_Link" size="35" maxlength="150" value="<%=Ads_Link%>"></td>
                            <td>Liên kết:</td>
                            <td class="in-input">
                                <select name="Ads_url" size="1">
                                    <option value="0" selected>Trang hiện tại</option>
                                    <option value="1">Mở tab mới</option>
                                </select>
                            </td>
                        </tr>
                    </table>
                </td>
                <td style="color: #F00;">*</td>
            </tr>
            <tr>
                <td>Hình Ảnh:</td>
                <td>
                    <table class="tb-it">
                        <tr>
                            <td>
                                <input name="Ads_ImagesPath" type="file" id="Ads_ImagesPath"><font color="#FF0000" size="1" face="Arial, Helvetica, sans-serif"><%=sAsd_ImagesPath%></font></td>
                            <td>Chiều cao:</td>
                            <td>
                                <input name="Ads_height" type="text" id="Ads_height" size="2" maxlength="3" value="<%=Ads_height%>"></td>
                            <td>Chiều rộng:</td>
                            <td>
                                <input name="Ads_width" type="text" id="Ads_width" size="2" maxlength="3" value="<%=Ads_width%>"></td>
                        </tr>
                    </table>
                </td>
                <td style="color: #F00;">*</td>
            </tr>
            <tr>
                <td>Vị trí:</td>
                <td>
                    <table class="tb-it">
                        <tr>
                            <td><%Call List_Ads_Position(Ads_Position,"----- Vị trí đặt quảng cáo -----")%></td>
                            <td><%Call List_Ads_StatusId(StatusId,"----- Trạng thái đăng -----")%></td>

                        </tr>
                    </table>
                </td>
                <td style="color: #F00;">*</td>
            </tr>

            <tr>
                <td>Chuyên mục:</td>
                <td>
                    <table class="tb-it">
                        <tr>
                            <td>
                                <select name="CategoryId" id="CategoryId">
                                    <option value="-2">-----Lựa chọn-----</option>
                                    <option value="0" <%if CategoryId=0 then%> selected<%end if%> style="color: Blue; background-color: #FFFFFF">Tất cả các chuyên mục</option>
                                    <option value="-1" <%if CategoryId=-1 then%> selected<%end if%> style="color: #000000; background-color: #E6E8E9">Nhiều chuyên mục (chọn sau)</option>
                                    <%Call List_CategoryWithoutSelectTag(CategoryId)%>
                                </select></td>
                            <td>
                                <input type="checkbox" name="Ads_OnlineChildren" value="1" <%if Ads_OnlineChildren=1 then%> checked<%end if%>></td>
                            <td style="width: 378px;">Bao gồm cả các chuyên mục con</td>

                        </tr>
                    </table>
                </td>
                <td style="color: #F00;">*</td>
            </tr>

            <tr>
                <td>Ghi chú:</td>
                <td>

                    <table class="tb-it">
                        <tr>
                            <td>
                                <input name="Ads_Note" type="text" id="VoteNote2" size="35" maxlength="100" value="<%=Ads_Note%>"></td>
                            <td class="">Mã màu:</td>
                            <td class="in-input">
                                <input name="idcolor_tex2" type="text" id="idcolor_tex2"  maxlength="20" value="" class=""></td>

                        </tr>
                    </table>
                </td>

                <td style="color: #F00;">*</td>
            </tr>
            <tr>
                <td></td>
                <td>

                    <table class="in-input">
                        <tr>
                            <td>
                                <input type="submit" name="Submit" value="Tạo mới"></td>
                            <td>
                                <input type="button" name="Submit2" value="Đóng cửa sổ" onclick="javascript: window.close();">
                            </td>

                        </tr>
                    </table>
                </td>

                <td style="color: #F00;"></td>
            </tr>

        </table>
    </form>
</body>
</html>
