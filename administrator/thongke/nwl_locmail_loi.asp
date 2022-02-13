<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/config.asp" -->

Bat dau Lan 4: <br>
<%
server.ScriptTimeout=server.ScriptTimeout*1000

		set rs=server.CreateObject("ADODB.Recordset")
		sql="select * from nwl_lan4 "
		rs.open sql,con,1
		i=1
		Tontai=0
		Khongco=0
		
		Do while not rs.EOF 
		Response.Write(i&"."&rs("email")&"<br>")
		
		'rs.MoveNext
		'Loop
		'rs.close
		'set rs=nothing
		'server.ScriptTimeout=server.ScriptTimeout/500				
		'Response.End


			' Xoa thanh vien 
			set rs_Sub=server.CreateObject("ADODB.Recordset")
			sql_Sub="select Top 1 * from nwl_register where nwl_Email='" & rs("email") & "'"
			rs_Sub.open sql_Sub,con,1
			if not rs_Sub.EOF then
			Tontai=Tontai+1
			Email=rs_Sub("nwl_Email")
			id=	rs_Sub("nwl_id")		
			
				sql_del_TV="delete nwl_register where nwl_Email='" & Email & "'"
				Set rs_del_TV=server.CreateObject("ADODB.Recordset")
				rs_del_TV.open sql_del_TV,con,1		
				set rs_del_TV=nothing
					Response.Write("Xoa thanh vien<br>")
				
				
				sql_del_TV="delete [nwl_BanTin] where nwl_Email_id='" & id & "'"
				Set rs_del_TV=server.CreateObject("ADODB.Recordset")
				rs_del_TV.open sql_del_TV,con,1		
				set rs_del_TV=nothing
					Response.Write("Xoa Ban tin<br>")			
			
			else
			
			End if
			Khongco=Khongco+1
			rs_Sub.close
			set rs_Sub=nothing
		i=i+1
		rs.MoveNext
		Loop
		rs.close
		set rs=nothing
%>
<br>
 Ket thuc: 
 Khongco:<%=Khongco%><br>
 Tontai:<%=Tontai%>
 <%
 server.ScriptTimeout=server.ScriptTimeout/1000
 %>