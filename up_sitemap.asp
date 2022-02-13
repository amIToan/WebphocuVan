<!--#include virtual="/include/config.asp"-->
<!--#include virtual="/include/Fs_liblary.asp"-->
<!--#include virtual="/include/lib_ajax.asp"-->
<%    
    Response.Clear
    Response.Buffer = True
    Response.AddHeader "Connection", "Keep-Alive"
    Response.CacheControl = "public"
    
    Dim strFolderArray, lngFolderArray
    Dim strUrlRoot, strPhysicalRoot, strFormat
    Dim strUrlRelative, strExt

    Dim objFSO, objFolder, objFile

    strPhysicalRoot = Server.MapPath("/")
    Set objFSO = Server.CreateObject("Scripting.Filesystemobject")
    
    strUrlRoot = "http://" & Request.ServerVariables("HTTP_HOST")
    
    sql_xmL_cate="SELECT CategoryID,CategoryName,ParentCategoryID,CategoryLink,CategoryLevel,CategoryOrder,"
    sql_xmL_cate=sql_xmL_cate&" CategoryHome,CategoryStatus,CategoryLoai,LanguageId FROM NewsCategory WHERE LanguageId='VN' ORDER BY CategoryOrder"

    Set rs_XML_cate=Server.CreateObject("ADODB.Recordset")
	rs_XML_cate.open sql_xmL_cate,con,1

    IF  not rs_XML_cate.eof THEN
            'thuc hien tao dom va lap duong link
            Dim theDom, theRoot, theParent,theChild, theID, docInstruction
            const NODE_ELEMENT = 1
            const xmlns = "http://www.sitemaps.org/schemas/sitemap/0.9"

            Set theDom = Server.CreateObject("Microsoft.XMLDOM")
            Set theRoot = theDom.createElement("urlset")
            Set theID = theDom.createAttribute("xmlns")
            theID.Text = xmlns
            theRoot.setAttributeNode theID
            theDom.appendChild theRoot

            'thuc hien in du lieu ra ngoài man hinh
            response.write("<link rel='stylesheet' href='https://www.w3schools.com/w3css/4/w3.css'>")

            response.write("<table class='w3-table-all w3-table'>")
            response.write("<tr><th>STT</th><th>Link</th><th>LastEditedDate</th><th>Creator</th></tr>")           
            i=1
        do while not rs_XML_cate.eof 
            CategoryId      =   rs_XML_cate("CategoryId")
            CategoryName    =   rs_XML_cate("CategoryName")
            CategoryLink    =   rs_XML_cate("CategoryLink") 
            CategoryLoai    =   rs_XML_cate("CategoryLoai")            
            
        IF CategoryLoai<> 7 THEN
            if CategoryLink<>"" then
                if CategoryLink="/" then
                    link_location=strUrlRoot
                else
                    link_location=CategoryLink         
                end if
            else
                link_location=strUrlRoot&"/"&CategoryId&"/"&Replace(Uni2NONE(CategoryName)," ","-") 
            end if

            Set theParent = theDom.createNode(NODE_ELEMENT, "url", xmlns)     
            Set theChild = theDom.createNode(NODE_ELEMENT, "loc", xmlns)
            theChild.Text = Server.HtmlEncode(link_location)
            theRoot.appendChild theParent
            theParent.appendChild theChild

            Set theChild = theDom.createNode(NODE_ELEMENT, "priority", xmlns)
            theChild.Text = "1.0"
            theRoot.appendChild theParent
            theParent.appendChild theChild

            Set theChild = theDom.createNode(NODE_ELEMENT, "changefreq", xmlns)
            theChild.Text = "weekly"   
            theRoot.appendChild theParent
            theParent.appendChild theChild   
            response.write("<tr><td>"&i&"</td><td>"&Server.HtmlEncode(link_location)&"<td><td>&nbsp&nbsp<td><td>&nbsp&nbsp<td></tr>")
    
            'thuc hien tao danh sach link url doi voi cac bai viet
            sql_xml="SELECT CategoryID,ParentCategoryID,CategoryHome,CategoryStatus,CategoryLoai,NewsID,"
            sql_xml=sql_xml&" Title,LastEditedDate,Creator FROM V_NEWS WHERE CategoryID='"&CategoryId&"'  order by LastEditedDate desc"

            Set rs_XML=Server.CreateObject("ADODB.Recordset")
	        rs_XML.open sql_xml,con,1

            IF  not rs_XML.eof THEN
                do while not rs_XML.eof 
                    CategoryId      =   rs_XML("CategoryId")
                    categoryloai    =   rs_XML("CategoryLoai")
                    NewsID          =   rs_XML("NewsId")
                    Title           =   rs_XML("Title")
                    LastEditedDate  =   rs_XML("LastEditedDate")
                    Creator         =   rs_XML("Creator")
 
                    'link_location=strUrlRoot&"/"&CategoryId&"/"&NewsID&"/"&Replace(Uni2NONE(Title)," ","-")&".html"
					link_location=strUrlRoot&func_uri(CategoryId,NewsID,Title)
					

                    Set theParent = theDom.createNode(NODE_ELEMENT, "url", xmlns)     
                    Set theChild = theDom.createNode(NODE_ELEMENT, "loc", xmlns)
                    theChild.Text = Server.HtmlEncode(link_location)
                    theRoot.appendChild theParent
                    theParent.appendChild theChild

                    Set theChild = theDom.createNode(NODE_ELEMENT, "priority", xmlns)
                    theChild.Text = "1.0"
                    theRoot.appendChild theParent
                    theParent.appendChild theChild

                    Set theChild = theDom.createNode(NODE_ELEMENT, "changefreq", xmlns)
                    theChild.Text = "weekly"   
                    theRoot.appendChild theParent
                    theParent.appendChild theChild  
    
                    'in ra ngoai man hinh du lieu   
                    response.write("<tr><td>"&i&"</td><td>"&Server.HtmlEncode(link_location)&"</td><td>"&LastEditedDate&"</td><td>"&Creator&"</td></tr>")  
                i=i+1   
                rs_XML.MoveNext()
                loop  
                'CAP NHAT LINK DOI VOI CAC CHUYEN MUC 
            END IF'end if rs_XML 
        
        END IF 'END IF CATEGORYLOAI 
        i=i+1    
        rs_XML_cate.MoveNext()
        loop           
    END IF'end if rs_XML_cate
    response.write("</table>")
    Set docInstruction = theDom.createProcessingInstruction("xml","version='1.0' encoding='UTF-8'")
    theDom.insertBefore docInstruction, theDom.childNodes(0)
    theDom.Save Server.MapPath("sitemap.xml")       
%>