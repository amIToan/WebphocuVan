<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%Sub write_header_av(Catid,lang)%>
  <div align="center">
  <center>
  <table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="775" id="AutoNumber16">
    <tr>
      <td width="100%">
      <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" bordercolor="#000000" id="table119" bgcolor="#FFFFFF">
		<tr>
			<td>
      <object classid="clsid:D27CDB6E-AE6D-11CF-96B8-444553540000" id="ShockwaveFlash4" width="775" height="84">
      <param name="Movie" value>
      <param name="Src" value="../../images/banner01.swf">
      <param name="Play" value="-1">
      <param name="Loop" value="-1">
      <param name="Quality" value="Best">
      <param name="SAlign" value>
      <param name="Menu" value="-1">
      <param name="Base" value>
      <param name="Scale" value="ExactFit">
      <param name="DeviceFont" value="0">
      <param name="EmbedMovie" value="-1">
      <param name="SWRemote" value>
                                      <embed SRC="../../images/banner01.swf" width="775" height="84" PLAY="true" LOOP="true" QUALITY="high" PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"></embed></object>
      
      </td>
		</tr>
		</table>
		</td>
    </tr>
    <tr>
      <td width="100%">
      <table border="0" cellpadding="0" style="border-collapse: collapse" width="775" bordercolor="#000000" id="table46">
		<tr>
			<td background=/home_01a_1.gif" height="20" width="20%">
			<p align="left">
			<font face="Tahoma" style="font-size: 8pt"><%=GetNameOfWeekDay(now)%>, <%=GetFullDate(now)%></font></td>
			<td background=/home_01a_1.gif" height="20">
			<p align="right" style="margin-right: 10px">
			</td>
		</tr>
		</table>
		</td>
    </tr>
  </table>
  </center>
</div>
<%End sub%>


<%Sub write_footer_av(CatId,lang)%>
  <table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="775" id="table72">
    <tr>
      <td bgcolor="#FFFFFF">
      <img border="0" src=/spacer.gif" width="1" height="1"></td>
    </tr>
    <tr>
      <td bgcolor="#F9F9F9">
      <p align="center" style="margin-top: 7px; margin-bottom: 7px">
      <font face="Verdana" style="font-size: 8pt" color="#808080"><b>CÔNG TY CỔ 
		PHẦN THƯƠNG MẠI 
		VÀ PHÁT TRIỂN CÔNG NGHỆ CAO</b></font></td>
    </tr>
    </table>
  

<%End sub%>

<%Function IsExistAv(AVid,Av_Title,Av_Path)
	'Truyền vào AVid
	'Trả về:	Av_Title: Av_Title
	'			 Av_Path: Av_Path
	'	 	   IsExistAv: =true: Tồn tại Audio - Video clips;
	'					  =false: Không tồn tại
	sqlAv="select Av_Title,Av_Path from AudioVideo where Av_id=" & AVid
	Dim rsAv
	set rsAv=server.CreateObject("ADODB.Recordset")
	rsAv.open sqlAv,con,3
	
	if rsAv.eof then
		rsAv.close
		set rsAv=nothing
		IsExistAv=false
		Exit Function
	end if
	Av_Title=rsAv("Av_Title")
	Av_Path=rsAv("Av_Path")
	rsAv.close
	set rsAv=nothing
	IsExistAv=true
End Function%>

<%Sub UpdateAVCounter(Avid)
	'Kiểm tra Cookies để chống refresh
	Cookies_Avid=Request.Cookies("Cookies_Avid")
	if IsNumeric(Cookies_Avid) then
		Cookies_Avid=Clng(Cookies_Avid)
	else
		Cookies_Avid=0
	end if
	if Cookies_Avid<>Avid then
		'Nếu hợp lệ Gọi UpdateAvCounter
	    Dim cm
    	set cm = CreateObject("ADODB.Command")
    	cm.ActiveConnection = strConnString

		cm.commandtype=4 'adstoredProc
		cm.CommandText = "UpdateAvCounter"
		cm.Parameters.Append cm.CreateParameter("IPAddress", 200, 1, 15, Request.ServerVariables("REMOTE_ADDR"))
		cm.Parameters.Append cm.CreateParameter("AvId", 3, 1, 4, AvId)
		'Set objparameter=objcommand.CreateParameter (name,type,direction,size,value)
		'cm.Parameters.Append cm.CreateParameter("counter", 3, 2, 4)
		cm.execute()
		set cm=nothing
		'Ghi lại Cookies
		Response.Cookies("Cookies_Avid")=Avid
		Response.Cookies("Cookies_Avid").Expires=DateAdd("n",5,now()) 'Chỉ lưu Cookies trong 5 phút
	end if
End Sub%>