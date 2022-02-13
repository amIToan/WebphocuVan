<%@  language="VBSCRIPT" codepage="65001" %>
<%
    Set Upload = Server.CreateObject("Persits.Upload")
	Upload.OverwriteFiles = True
    Path = Server.MapPath("\administrator\images\img_nv\")
    Upload.Save 
      On Error Resume Next
      For Each File in Upload.Files
             filename = File.FileName
              Set Jpeg = Server.CreateObject("Persits.Jpeg")             
              Jpeg.OpenBinary( File.Binary )
              Jpeg.PreserveAspectRatio  = True
        Jpeg.Width = 640
        Jpeg.ResolutionX = 72
        Jpeg.ResolutionY = 72
        Path = Server.MapPath("/administrator/images/img_nv/"&filename)       
        Jpeg.Save Server.MapPath("/administrator/images/img_nv/"&filename)           
      Next 
%>