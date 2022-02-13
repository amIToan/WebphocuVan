<!--#include virtual="/include/config.asp" -->
<!--#include virtual="/administrator/inc/Fs_liblary.asp" -->
<%


    Set Upload = Server.CreateObject("Persits.Upload")
    Upload.SetMaxSize 10000000, True 'Dat kich co upload la` 1MB
    Upload.codepage=65001
    Upload.Save   
    key_  =  Upload.Form("_key")
    stt_status = 0
    IF key_ <> "" And  key_ = "Add" THEN
       'Get date
        'INSERT &  UPLOAD IMG ----------------------------------------------
        Path=server.MapPath("/images/logo")   
        set FIcon = Upload.Files("icon")
        IF FIcon Is Nothing THEN
        	FIcon_=""
        ELSE
           Filetype = Right(FIcon.FileName,len(FIcon.Filename)-Instr(FIcon.Filename,"."))
           IF Lcase(Filetype)<>"ico" THEN
        		FIcon_=""
           ELSE
               dt = Trim(Replace(getDateServer(),"/",""))
               dt = Trim(Replace(dt,":",""))
               dt = Trim(Replace(dt," ",""))             
        	   FIcon_="IC-"&dt&"1."&Filetype
        	   FIcon.SaveAs Path &"\"&FIcon_            
            END IF
        END IF
        
        'INSERT &  UPLOAD IMG ----------------------------------------------
        Path=server.MapPath("/images/logo")
        set FLogo = Upload.Files("icon")
        IF FLogo Is Nothing THEN
        	FLogo_=""
        ELSE
           Filetype = Right(FLogo.FileName,len(FLogo.Filename)-Instr(FLogo.Filename,"."))
           IF Lcase(Filetype)<>"jpg" and Lcase(Filetype)<>"gif" and Lcase(Filetype)<>"jpeg" and Lcase(Filetype)<>"png" and Lcase(Filetype)<>"ico" THEN
        		FLogo_=""
           ELSE
               dt = Trim(Replace(getDateServer(),"/",""))
               dt = Trim(Replace(dt,":",""))
               dt = Trim(Replace(dt," ",""))             
        	   FLogo_="IMG-L"&dt&"0."&Filetype
        	   FLogo.SaveAs Path &"\"&FLogo_            
            END IF
        END IF
        
        'INSERT &  UPLOAD IMG ----------------------------------------------
        Path=server.MapPath("/images/logo")
        set FLgF = Upload.Files("logoF")
        IF FLgF Is Nothing THEN
        	FLgF_=""
        ELSE
           Filetype = Right(FLgF.FileName,len(FLgF.Filename)-Instr(FLgF.Filename,"."))
           IF Lcase(Filetype)<>"jpg" and Lcase(Filetype)<>"gif" and Lcase(Filetype)<>"jpeg" and Lcase(Filetype)<>"png" and Lcase(Filetype)<>"ico" THEN
        		FLgF_=""
           ELSE
               dt = Trim(Replace(getDateServer(),"/",""))
               dt = Trim(Replace(dt,":",""))
               dt = Trim(Replace(dt," ",""))             
        	   FLgF_="IMG-LF"&dt&"."&Filetype
        
        	   FLgF.SaveAs Path &"\"&FLgF_            
            END IF
        END IF
         
       cs_Name         = Upload.Form("company")
       cs_address      = Upload.Form("address")
       cs_Tel          = Upload.Form("Tel")
       cs_Hotline      = Upload.Form("Hotline")
       cs_Email        = Upload.Form("Email")
       cs_Website      = Upload.Form("Website")
       cs_IdTax        = Upload.Form("masothue")
       cs_GPKD         = Upload.Form("GPKD")
       cs_meta_title   = Upload.Form("page_title")
       cs_meta_desc    = Upload.Form("meta_description")
       cs_meta_keyword = Upload.Form("meta_keywords")
       cs_calltime     = Upload.Form("calltime")
       cs_idgoogle     = Upload.Form("idgoogle")
       cs_idyoutube    = Upload.Form("idyoutube")
       cs_Sky          = Upload.Form("idskype")
       cs_idgplus      = Upload.Form("idgplus")
      
       cs_idface       = Upload.Form("idfacebook")
       cs_intro        = Upload.Form("csDesc") 
       cs_idtwiter     = Upload.Form("idtwiter")
       cs_ckeckmain    = Upload.Form("ckeckmain")
        
       IF cs_ckeckmain <> "" AND  IsNumeric(cs_ckeckmain) THEN
            cs_ckeck_ = 1
       ELSE
            cs_ckeck_ = 0          
       END  IF

    
       cs_idLang       = "VN"
       
       sql = "INSERT INTO Company ( [company],[Tel],[calltime],[Hotline],[Email],[Website],[address],[Masothue],[GPKD],[page_title],[meta_description],[meta_keywords],[icon],[Logo],[LogoF],[idgoogle],[idtwiter],[introduction],[idgplus],[idyoutube],[idskype],[idfacebook],[lang],[show] )"
       sql = sql&" VALUES ( "
       sql = sql&"  N'"&cs_Name&"'"           ' Name
       sql = sql&" ,N'"&cs_Tel&"'"            ' Tel
       sql = sql&" ,N'"&cs_calltime&"'"       ' Call time
       sql = sql&" ,N'"&cs_Hotline&"'"        ' Hotline
       sql = sql&" ,N'"&cs_Email&"'"          ' Email
       sql = sql&" ,N'"&cs_Website&"'"        ' Website
       sql = sql&" ,N'"&cs_address&"'"        ' Address
       sql = sql&" ,N'"&cs_IdTax&"'"          ' ID Tax
       sql = sql&" ,N'"&cs_GPKD&"'"           ' GPKD
       sql = sql&" ,N'"&cs_meta_title&"'"     ' SEO 
       sql = sql&" ,N'"&cs_meta_desc&"'"      ' SEO 
       sql = sql&" ,N'"&cs_meta_keyword&"'"   ' SEO 
       sql = sql&" ,N'"&FIcon_&"'"            ' File ico
       sql = sql&" ,N'"&FLogo_&"'"            ' File Logo
       sql = sql&" ,N'"&FLgF_&"'"             ' File logo
       sql = sql&" ,N'"&cs_idgoogle&"'"       ' ID Google Analytic
       sql = sql&" ,N'"&cs_idtwiter&"'"       ' idtwiter
       sql = sql&" ,N'"&cs_intro&"'"          ' Introduct  
       sql = sql&" ,N'"&cs_idgplus&"'"        ' G+ link
       sql = sql&" ,N'"&cs_idyoutube&"'"      ' Youtobe link
       sql = sql&" ,N'"&cs_Sky&"'"            ' Sky
       sql = sql&" ,N'"&cs_idface&"'"         ' Id Face
       sql = sql&" ,N'"&cs_idLang&"'"         ' Lang
       sql = sql&" , '"&cs_ckeck_&"'"                     ' show
       sql = sql&" ) "                        ' END
        
       on error resume next
       con.Execute sql,recaffected
       if err<>0 then
         'not ok
          stt_status = 0
       else
         'ok
          stt_status = 1
       end if
       conn.close
      Response.Write stt_status

    ELSEIF key_ <> "" And  key_ = "Update" THEN
       cs_id_          = Upload.Form("_id")
       cs_Name         = Upload.Form("FEdit_copany")
       cs_Address      = Upload.Form("address")
       cs_Tel          = Upload.Form("Tel")
       cs_Hotline      = Upload.Form("Hotline")
       cs_Email        = Upload.Form("Email")
       cs_Website      = Upload.Form("Website")
       cs_IdTax        = Upload.Form("masothue")
       cs_GPKD         = Upload.Form("GPKD")
       cs_meta_title   = Upload.Form("page_title")
       cs_meta_desc    = Upload.Form("meta_description")
       cs_meta_keyword = Upload.Form("meta_keywords")
       cs_calltime     = Upload.Form("calltime")
       cs_idgoogle     = Upload.Form("idgoogle")
       cs_idyoutube    = Upload.Form("idyoutube")
       cs_Sky          = Upload.Form("idskype")
       cs_idgplus      = Upload.Form("idgplus")
       cs_idface       = Upload.Form("idfacebook")
       cs_Ficon        = Upload.Form("_Ficon")
       cs_FLogo        = Upload.Form("_FLogo")
       cs_FlogoF       = Upload.Form("_FlogoF")
       cs_intro        = Upload.Form("csDesc") 
       cs_idtwiter     = Upload.Form("idtwiter")
       cs_ckeckmain     = Upload.Form("ckeckmain")

      IF cs_ckeckmain <> "" AND  IsNumeric(cs_ckeckmain) THEN
            cs_ckeck_ = 1
       ELSE
            cs_ckeck_ = 0          
       END  IF

       cs_idLang       = "VN"

       sql_ = "UPDATE [Company] SET "
       sql_ = sql_&" [company] = N'"&cs_Name&"'"
       sql_ = sql_&",[Tel] = N'"&cs_Tel&"'"
       sql_ = sql_&",[calltime] = N'"&cs_calltime&"'"
       sql_ = sql_&",[Hotline] = N'"&cs_Hotline&"'"
       sql_ = sql_&",[Email] = N'"&cs_Email&"'"
       sql_ = sql_&",[Website] = N'"&cs_Website&"'"
       sql_ = sql_&",[address] = N'"&cs_Address&"'"
       sql_ = sql_&",[Masothue] = N'"&cs_IdTax&"'"
       sql_ = sql_&",[GPKD] = N'"&cs_GPKD&"'"
       sql_ = sql_&",[page_title] = N'"&cs_meta_title&"'"
       sql_ = sql_&",[meta_description] = N'"&cs_meta_desc&"'"
       sql_ = sql_&",[meta_keywords] = N'"&cs_meta_keyword&"'"
       sql_ = sql_&",[idgoogle] = N'"&cs_idgoogle&"'"
       sql_ = sql_&",[introduction] = N'"&cs_intro&"'"
       sql_ = sql_&",[idgplus] = N'"&cs_idgplus&"'"
       sql_ = sql_&",[idyoutube] = N'"&cs_idyoutube&"'"
       sql_ = sql_&",[idskype] = N'"&cs_Sky&"'"
       sql_ = sql_&",[idfacebook] = N'"&cs_idface&"'"
       sql_ = sql_&",[idtwiter] = N'"&cs_idtwiter&"'"
       sql_ = sql_&",[lang] = N'"&cs_idLang&"'"
       sql_ = sql_&",[show] = '"&cs_ckeck_&"'"
      
    'INSERT &  UPLOAD IMG ----------------------------------------------
           
        Path=server.MapPath("/images/logo")   
        set FIcon = Upload.Files("icon")
        IF FIcon Is Nothing THEN
        	FIcon_=""
        ELSE
           Filetype = Right(FIcon.FileName,len(FIcon.Filename)-Instr(FIcon.Filename,"."))
           IF Lcase(Filetype)<>"ico" THEN
        		FIcon_=""
           ELSE
               dt = Trim(Replace(getDateServer(),"/",""))
               dt = Trim(Replace(dt,":",""))
               dt = Trim(Replace(dt," ",""))             
        	   FIcon_="IC-"&dt&"1."&Filetype
        	   FIcon.SaveAs Path &"\"&FIcon_  
               sql_= sql_&",[icon] = N'"&FIcon_&"'"
               Url1 = "/images/logo/"&cs_Ficon
               DelFile(Url1)        
            END IF
        END IF
        
        'INSERT &  UPLOAD IMG ----------------------------------------------
        Path=server.MapPath("/images/logo")
        set FLogo = Upload.Files("Logo")
        IF FLogo Is Nothing THEN
        	FLogo_=""
        ELSE
           Filetype = Right(FLogo.FileName,len(FLogo.Filename)-Instr(FLogo.Filename,"."))
           IF Lcase(Filetype)<>"jpg" and Lcase(Filetype)<>"gif" and Lcase(Filetype)<>"jpeg" and Lcase(Filetype)<>"png" and Lcase(Filetype)<>"ico" THEN
        		FLogo_=""
           ELSE
               dt = Trim(Replace(getDateServer(),"/",""))
               dt = Trim(Replace(dt,":",""))
               dt = Trim(Replace(dt," ",""))             
        	   FLogo_="IMG-L"&dt&"0."&Filetype
        	   FLogo.SaveAs Path &"\"&FLogo_  
               sql_ = sql_& ",[Logo] = N'"&FLogo_&"'"
               Url2 = "/images/logo/"&cs_FLogo
               DelFile(Url2) 
            END IF
        END IF
        
        'INSERT &  UPLOAD IMG ----------------------------------------------
        Path=server.MapPath("/images/logo")
        set FLgF = Upload.Files("logoF")
        IF FLgF Is Nothing THEN
        	FLgF_=""
        ELSE
           Filetype = Right(FLgF.FileName,len(FLgF.Filename)-Instr(FLgF.Filename,"."))
           IF Lcase(Filetype)<>"jpg" and Lcase(Filetype)<>"gif" and Lcase(Filetype)<>"jpeg" and Lcase(Filetype)<>"png" and Lcase(Filetype)<>"ico" THEN
        		FLgF_=""
           ELSE
               dt = Trim(Replace(getDateServer(),"/",""))
               dt = Trim(Replace(dt,":",""))
               dt = Trim(Replace(dt," ",""))             
        	   FLgF_="IMG-LF"&dt&"."&Filetype         
        	   FLgF.SaveAs Path &"\"&FLgF_ 
               sql_ = sql_& ",[LogoF] = N'"&FLgF_&"'"
               Url3 = "/images/logo/"&cs_FLogoF
               DelFile(Url3) 
            END IF
        END IF

     sql_= sql_&" WHERE  ID = '"&cs_id_&"'"


      on error resume next
      con.Execute sql_,recaffected
      if err<>0 then
        'not ok
         stt_status = 0
      else
        'ok
         stt_status = 1
      end if
      conn.close
       Response.Write stt_status   
    END IF



%>