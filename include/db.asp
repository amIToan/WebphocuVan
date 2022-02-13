<%
'	set connectDB = CreateObject("ADODB.Connection")
'		connectDB.ConnectionString = "Provider=SQLOLEDB; data source=IT4;database=xsoft; User ID=sa; password=1234567;"
'	    '"Provider=SQLOLEDB;Data Source=serverName;" Initial Catalog=databaseName;User ID=MyUserID;Password=MyPassword;"
'	connectDB.Open
'	if IsObject(connectDB) then
'		if connectDB.State = 1 then

			'Response.Redirect("info.asp?var1=1&va2=2")	
'		else 
'			Response.Redirect("errors.asp?var1=1&va2=2")
'		end if
'		else 
'			Response.Redirect("errors.asp")
'	end if
%>
<% 'Response.Redirect Server.HTMLEncode("newpage.asp?var1=5&var2=7") %>

<%

Dim con
'Function connect db
Sub OpenConnect()
	set con = CreateObject("ADODB.Connection")
			con.ConnectionString = "Provider=SQLOLEDB; data source=VIETSOFT\SQLEXPRESS;database=xseodb; User ID=sa; password=1234567;"
		con.Open
		if IsObject(con) then
			if con.State = 1 then
				'Response.Write("<script> alert('Connected database !') ;</script>")
				'Response.Redirect("info.asp?var1=1&va2=2")	
			else 
				'Response.Redirect("errors.asp?var1=1&va2=2")
			end if
			else 
				'Response.Redirect("errors.asp")
		end if
End sub

'Hủy kết nối CSDL

Sub DestroyConnect()
	con.Close 
	set con = nothing
End sub




'"Provider=SQLOLEDB;Data Source=serverName;" Initial Catalog=databaseName;User ID=MyUserID;Password=MyPassword;"
%>	

		