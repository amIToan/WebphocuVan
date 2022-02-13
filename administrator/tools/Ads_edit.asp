<%@  language="VBSCRIPT" codepage="65001" %>
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<!--#include virtual="/administrator/inc/func_cat.asp" -->
<!--#include virtual="/administrator/inc/func_Ads.asp" -->
<%
f_permission = administrator(false,session("user"),"m_ads")
if f_permission <= 1 then
	response.Redirect("/administrator/info.asp")
end if
%>
<%
	Call AuthenticateWithRole(AdvertisementCategoryId,Session("LstRole"),"ap")
	Ads_id=GetNumeric(Request.querystring("param"),0)
	Dim rs
	if Request.Querystring("action")="Insert" then
		sError=False
		Set Upload = Server.CreateObject("Persits.Upload")
		
		Upload.SetMaxSize 10000000, True 'Dat kich co upload la` 1MB
		Upload.codepage=65001
		Upload.Save
		'Ads_Title
		Ads_Title=ReplaceHTMLToText(Upload.form("Ads_Title"))
		if Ads_Title="" then
			sError=True
		end if
		'Ads_Link
		Ads_Link=ReplaceHTMLToText(Upload.form("Ads_Link"))
		Ads_url=Clng(Upload.form("Ads_url"))
		Ads_Target=ReplaceHTMLToText(Upload.form("Ads_Target"))



	
		idcolor1 =trim(Upload.form("idcolor_tex1"))
		idcolor2=trim(Upload.form("idcolor_tex2"))

		Ads_width=trim(Upload.form("Ads_width"))
		Ads_height=trim(Upload.form("Ads_height"))
        
        IF Not IsNumeric(Ads_width) THEN  
            Awidth = ""
        ELSE
            Awidth = Ads_width
        END IF
        IF Not IsNumeric(Ads_height) THEN  
            Aheight = ""
        ELSE
            Aheight = Ads_height
        END IF

		
		'Ads_Position
		Ads_Position=GetNumeric(Upload.form("Ads_Position"),-1)
		if Clng(Ads_Position)=-1 then
			sError=True
		end if
		'StatusId
		StatusId=ReplaceHTMLToText(Upload.form("StatusId"))
		if statusId<>"eded" and statusId<>"apap" then
			sError=True
		end if
		'CategoryId
		CategoryId=GetNumeric(Upload.form("CategoryId"),-2)

		if CategoryId=-3 then
			sError=True
		end if
		'Ads_OnlineChildren
		Ads_OnlineChildren=GetNumeric(Upload.form("Ads_OnlineChildren"),0)
		'Ads_Note
		Ads_Note=ReplaceHTMLToText(Upload.form("Ads_Note"))
		
		if not sError then
			set rs=server.CreateObject("ADODB.Recordset")



			

            Ads_Type = 1





			sql=" Update Ads set " 
            sql=sql&"Ads_Title=N'" & Ads_Title & "'"
			sql=sql&",Ads_Link='" & Ads_Link & "'"  
			sql=sql&",Ads_Target='" & Ads_Target & "'"
        '---------------------------------------------------------------------------------------
            set FImg3 = Upload.Files("Ads_ImagesPath")
            If FImg3 Is Nothing Then
            	FImg3_=""
            else
               Filetype = Right(FImg3.FileName,len(FImg3.Filename)-Instr(FImg3.Filename,"."))
               	   if Lcase(Filetype)<>"jpg" and Lcase(Filetype)<>"gif" and Lcase(Filetype)<>"jpeg" and Lcase(Filetype)<>"png" and Lcase(Filetype)<>"ico" then
            		sError=True
            		FImg3_=""
               else
                   dt = Trim(Replace(getDateServer(),"/",""))
                   dt = Trim(Replace(dt,":",""))
                   dt = Trim(Replace(dt," ",""))             
            	   FImg3_="IMG_"&dt&"3."&Filetype
               end if
            End If
            
            if FImg3_ <>"" then 
                'del file                
                UriF = "/images_upload/"&oFilePic3&""                   
                DelFile(UriF)   
                'save new
                sql=sql&",Ads_ImagesPath='" & FImg3_ & "'"
            	FImg3.SaveAs Path &"\"&FImg3_
            end if
        '---------------------------------------------------------------------------------------
			sql=sql&" ,Ads_Type=" & Ads_Type 
			sql=sql&" ,Ads_width = '" & Awidth  & "'" 
			sql=sql&" ,Ads_height='" & Aheight & "'"   
			sql=sql&" ,StatusId='" & StatusId & "'" 
			sql=sql&" ,idcolor_tex1= N'" & idcolor1 & "'"
			sql=sql&" ,idcolor_tex2= N'" & idcolor2 & "'" 
			sql=sql&" ,Ads_LastEdited='" & now() & "'" 
			sql=sql&" ,Ads_Note=N'" & Ads_Note & "'"
			sql=sql&" ,Ads_url='" & Ads_url & "'"
			sql=sql&" Where Ads_id=" & Ads_id
				
			response.write sql & "<br>"
			'response.end
			rs.open sql,con,1


			if CategoryId=-1 then
			'Insert vào nhiều chuyên mục khác nhau
				set rs=nothing
				set upload=nothing
				response.redirect ("Ads_MultiCat.asp?id=" & Ads_id)
			else
				sql="Delete AdsDistribution where Ads_id=" & Ads_id & ";" &_
					"Insert Into AdsDistribution " &_
					"(Ads_id,CategoryId, Ads_OnlineChildren, Ads_Order) values " &_
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
		sql="SELECT	TOP 2 a.Ads_id, a.Ads_Title, a.Ads_Link, a.Ads_Target, a.Ads_ImagesPath, a.Ads_width, " &_
			"		a.Ads_height, a.Ads_Position, a.StatusId, a.Ads_Creator, a.Ads_Type, " &_
            "		a.Ads_CreationDate, a.Ads_LastEditor, a.Ads_LastEdited, a.Ads_Note, a.Ads_url, " &_
            "		d.Ads_OnlineChildren, d.CategoryId ,a.idcolor_tex1,a.idcolor_tex2 " &_
			" FROM   Ads a INNER JOIN " &_
            "			AdsDistribution d ON a.Ads_id = d.Ads_id " &_
			" WHERE     (a.Ads_id = " & Ads_id & ")"

		Set rs=server.CreateObject("ADODB.Recordset")
		rs.open sql,con,1

		Ads_Title=rs("Ads_Title")

		id_color1= rs("idcolor_tex1")
		id_color2= rs("idcolor_tex2")
		Ads_Link=rs("Ads_Link")
		Ads_Target=rs("Ads_Target")
		Ads_ImagesPath=Trim(rs("Ads_ImagesPath"))
		Ads_Type=Clng(rs("Ads_Type"))
		Ads_height=rs("Ads_height")
		Ads_width=rs("Ads_width")
		Ads_Position=rs("Ads_Position")
		StatusId=rs("StatusId")
		Ads_Note=rs("Ads_Note")
		CategoryId=rs("CategoryId")
		Ads_OnlineChildren=rs("Ads_OnlineChildren")
		Ads_url=rs("Ads_url")
		rs.movenext
		if not rs.eof then
			'Dấu hiệu để nhận biết, quảng cáo hiển thị ở nhiều chuyên mục
			CategoryId=-1
			Ads_OnlineChildren=0
		end if
		rs.close
		set rs=nothing
	end if 'Of if Request.Form("action")="Insert" then
%>
<html>
<head>
    <title><%=PAGE_TITLE%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <script language="JavaScript1.2" src="/administrator/inc/common.js"></script>
</head>
<body>
    <form action="<%=request.ServerVariables("SCRIPT_NAME")%>?action=Insert&param=<%=Ads_id%>" method="post" enctype="multipart/form-data" name="fInsertEvent">
        <table border="0" class=" Tb-input Tb-in ">
            <tr>
                <th colspan="2">Chỉnh sửa Banner , Icon quảng cáo <%=Ads_ImagesPath %></th>
            </tr>

            <tr>
                <th colspan="2">
                    <%if Ads_ImagesPath<>"" then%>
                    <img src="/images_upload/<%=Ads_ImagesPath %>"  style="width:800px;"/>
                    <%end if%>
                </th>
            </tr>


            <tr>
                <td style="width: 24%;">Tiêu đề:</td>
                <td style="width: 76%;">

                    <table class="tb-it">
                        <tr>
                            <td>
                                <input name="Ads_Title" type="text" id="Ads_Title" size="35" maxlength="100" value="<%=Ads_Title%>"></td>
                            <td class="">Mã màu:</td>
                            <td class="in-input">
                                <input name="idcolor_tex1" type="text" id="idcolor_tex1" size="20" maxlength="20" value="<%=id_color1%>"></td>

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
            <!--  <tr>
                <td></td>
                <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      	<%'if Ads_ImagesPath<>"" then%>
                    	<input type="checkbox" name="RemoveImg" id="RemoveImg" value="1"><font face="Arial, Verdana" size="2">Bỏ Banner, Icon</font>
                    <%'end if%>
                </td>
            </tr> -->


            <tr>
                <td>Hình Ảnh:</td>
                <td>
                    <table class="tb-it">
                        <tr>
                            <td>

                                <input name="Ads_ImagesPath" type="file" id="Ads_ImagesPath">
                            </td>
                            <td></td>
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
                            <td>Chiều Rộng:</td>
                            <td>
                                <input name="Ads_width" type="text" id="Ads_width" size="2" maxlength="4" value="<%=Ads_width%>"></td>
                            <td>Chiều Cao:</td>

                            <td>
                                <input name="Ads_height" type="text" id="Ads_height" size="2" maxlength="4" value="<%=Ads_height%>"></td>

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
                                    <option value="-3">-----Lựa chọn-----</option>
                                    <option value="-2" <%if CategoryId=-2 then%> selected <%end if%>>-Trang chủ</option>
                                    <option value="0" <%if CategoryId=0 then%> selected <%end if%>>-Tất cả</option>
                                    <option value="-1" <%if CategoryId=-1 then%> selected <%end if%> style="color: #000000; background-color: #E6E8E9">Chuyên mục</option>
                                    <%Call List_CategoryWithoutSelectTag(CategoryId)%>
                                </select></td>
                            <td>
                                <input type="checkbox" name="Ads_OnlineChildren" value="1" <%if Ads_OnlineChildren=1 then%> checked <%end if%>></td>
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
                                <input name="idcolor_tex2" type="text" id="idcolor_tex2" size="35" maxlength="20" value="<%=id_color2 %>"></td>

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
                                <input type="submit" name="Submit" value="Cập nhật"></td>
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
