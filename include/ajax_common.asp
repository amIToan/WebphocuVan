<%@  language="VBSCRIPT" codepage="65001" %>
<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/include/lib_ajax.asp" -->

<%
    keyid = Request.Form("key")
    Dim sysMsg
    sysMsg =""
    Select Case keyid       
    Case "send-contact"                     
          Call func_contact() 
    Case "send-email"                     
          Call func_email()
    Case "changeLang"  
            lang=Request.Form("lang")                   
            Call changeLang(lang) 
    Case "payingcart"   
        Call updatepayingCarts()
    End Select
%>

<%sub changeLang(lang)
    Session("Language") = lang
end sub %>

<%sub  func_email()

    vr_status = -1
    FName       = Request.Form("inputName")
    FEmail      = Request.Form("inputEmail")    
    Fid         =  3
    Show        =  0

    IF (FName <> "")  and  (FEmail  <> "") THEN 
      sql_Fa = "INSERT INTO Email (Ten,email)  VALUES ( "
      sql_Fa = sql_Fa&"N'"&FName&"',"
      sql_Fa = sql_Fa& "'"&FEmail&"'"
      sql_Fa = sql_Fa&")"
    'Response.write sql_Fa
      set rsFa=server.CreateObject("ADODB.Recordset")   
       on error resume next
        con.Execute sql_Fa
        if err<>0 then
          vr_status = "0"
        else
             vr_status = "1"        
        end if
      set rsFa=nothing
    ELSE 
       vr_status = "-1"
    END IF

    Response.Write vr_status
end sub
%>

<%sub  func_contact()

    vr_status = -1

    FName       = Request.Form("F_Name")
    FEmail      = Request.Form("F_Email")
    F_Title        = Request.Form("F_Title")
    F_Content        = Request.Form("F_Content")



    Fid         =  3
    Show        =  0

    IF (FName <> "")  and  (FEmail  <> "") and (F_Title  <> "") and (F_Content  <> "")  THEN 
      sql_Fa = "INSERT INTO Y_KIEN (hovaten,email,tieude,noidung,faq,show)  VALUES ( "
      sql_Fa = sql_Fa&"N'"&FName&"',"
      sql_Fa = sql_Fa& "'"&FEmail&"',"
      sql_Fa = sql_Fa& "N'"&F_Title&"',"
      sql_Fa = sql_Fa& "N'"&F_Content&"',"
      sql_Fa = sql_Fa&"'"&Fid&"',"
      sql_Fa = sql_Fa&"'"&Show&"'"
      sql_Fa = sql_Fa&")"
      set rsFa=server.CreateObject("ADODB.Recordset")   
       on error resume next
        con.Execute sql_Fa
        if err<>0 then
          vr_status = "0"
        else
             vr_status = "1"        
        end if
      set rsFa=nothing
    ELSE 
       vr_status = "-1"
    END IF

    Response.Write vr_status
end sub
%>



<%Function getColVal(tabName,colName,query) %>
<% 
    if tabName = "" or colName = "" or query = "" then
        getColVal = ""
    else
        sqlCV = "Select top 1 "&colName&" from "&tabName&" where "&query
        set rsCV = server.CreateObject("ADODB.RecordSet")
   ' Response.Write sqlCV
        rsCV.open sqlCV,con,1
            if not rsCV.EOF then
                getColVal = Trim(rsCV(colName))
            else
                getColVal = ""
            end if
        set rsCV = nothing
    end if
%>
<%End Function %>




<%
    sub faq_byID(faqID) 
	sqlFa="SELECT top 1 * from Y_KIEN where faq=1 and show <> '0' and ID='"&faqID&"'"
	Set rsFA=Server.CreateObject("ADODB.Recordset")
	rsFA.open sqlFa,con,3
	if not rsFA.eof then
		Name =Trim(rsFA("hovaten"))
		conten_qs =Trim(rsFA("noidung"))	
        obj_fa	=rsFA("tieude")
        anw_fa	=rsFA("Traloi")
        
%>
<div class="col-lg-12 ">
    <img src="/images/icons/announce1.gif" border="0" class="img-responsive" style="margin-left: 0px;" />
    <h4 class="h4-faq"><%=Name %></h4>
    <p style="padding: 20px 0px 50px 0px; background: #FDEBEB; padding: 10px;">
        <i class="faq-obj"><%=obj_fa %></i><br>
        <span class="faq-qs">
            <b><u class="faq-qs">Câu hỏi: </u></b><%=conten_qs %>
        </span>
        <br />
        <span class="faq-aw">
            <b><u class="faq-qs">Trả lời: </u></b><%=anw_fa %>
        </span>
    </p>
</div>
<%    	

    end if
    end sub 
%>
