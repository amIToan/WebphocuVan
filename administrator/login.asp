<!--#include virtual="/include/config.asp"-->
<!--#include virtual="/include/constant.asp"-->
<!--#include virtual="/administrator/inc/md5.asp"-->
<!--#include virtual="/administrator/inc/func_common.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%      
	user=replace(request.Form("uid"),"'","''")
	password=md5(replace(request.Form("pwd"),"'","''"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	'sql="SELECT * FROM [User] u WHERE (u.UserName = N'" & user & "') "
	sql="SELECT * FROM [User] WHERE (UserName = N'" & user & "') AND (UserPwd= N'" & password & "')"
	'response.write(md5("123456"))
	rs.open sql,con,1
	if not rs.eof then

  ' Response.Write getDateServer()  
		'Session("NhanVienID")=rs("IDNhanVien")
		session("user")	= user	
		Set rs1 = Server.CreateObject("ADODB.Recordset")
		sql="SELECT     Nhanvien.Ho_Ten, Nhanvien.CMT,Nhanvien.NhanVienID, Nhanvien.imgNV, PhongBan.Description as room, ChucDanh.Description as TitleUser, StaffContract.EndDate, StaffContract.isMember FROM Nhanvien INNER JOIN StaffContract ON Nhanvien.NhanVienID = StaffContract.NhanVienID INNER JOIN ChucDanh ON StaffContract.ChucdanhID = ChucDanh.ChucDanhID INNER JOIN PhongBan ON StaffContract.PhongID = PhongBan.PhongID WHERE Nhanvien.NhanVienID='"& rs("IDNhanVien") &"' and (DATEDIFF(dd,StaffContract.EndDate,'" & getDateServer() & "')<= 0) order by StaffContract.id desc"
	'response.write(sql)

    rs1.open sql,con,1				
		if not rs1.eof then
			Session("user")		=	user
			Session("fullname")	=	rs1("Ho_Ten")
			Session("staffimg")	=	rs1("imgNV")
			Session("TitleUser")=	rs1("TitleUser")
			Session("NhanVienID")=	rs1("NhanVienID")
			Session("room")		=	rs1("room")			
			session("LstCat")	=	"0"		'Tat ca ca'c chuyen muc
			session("LstRole")	=	"0ad"
			'Update LastLoginDate
			Set rsTemp = Server.CreateObject("ADODB.Recordset")
			sql="Update [User] set LastLoginDate='" & getDateServer() & "' where Username=N'" & user & "'"
			Response.Write(sql)
			'on error Resume next
			rsTemp.open sql,con,3
			set rsTemp=nothing
			
			'Response.End()
			call UserOperation(user,hour(now)&":"&Minute(now)&"phut : Dang nhap thanh cong")
			response.Redirect("welcome.asp")
		else
			'Username chua duoc cap quyen
			rs.close
			set rs=nothing
			call UserOperation(user,hour(now)&":"&Minute(now)&"phut : Dang nhap khong thanh cong")
			response.Redirect("default.asp?sError=Inactive")
		end if
		rs1.close
		
	else
		'Neu khong co' thi` bao loi sai Username or Password
		rs.close
		set rs=nothing
		response.Redirect("default.asp?sError=Invalid")
	end if
%>
