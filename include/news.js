// JavaScript Document
function winpopup(urlx,param,twidth,theight)
{
	var strurl= urlx + '?param=' + param;
	var tposx= (screen.width- twidth)/2
	var tposy= (screen.height- theight)/2;

	var newWin=window.open(strurl,"moi","toolbar=no,width="+ twidth+",height="+ theight+ ",directories=no,status=no,scrollbars=yes,resizable=no, menubar=no")
	newWin.moveTo(tposx,tposy);
	newWin.focus();
}

function winpopup2(urlx,param,twidth,theight)
{
	var strurl= urlx + '?param=' + param;
	var tposx= (screen.width- twidth)/2
	var tposy= (screen.height- theight)/2;

	var newWin2=window.open(strurl,"window2","toolbar=no,width="+ twidth+",height="+ theight+ ",directories=no,status=no,scrollbars=yes,resizable=no, menubar=no")
	newWin2.moveTo(tposx,tposy);
	newWin2.focus();
}

function winpopup_vn(urlx,param,twidth,theight)
{
	var strurl= urlx + '?param=' + param;
	var tposx= (screen.width- twidth)/2
	var tposy= (screen.height- theight)/2;

	var newWin2=window.open(strurl,"window2","toolbar=no,width="+ twidth+",height="+ theight+ ",directories=no,status=no,scrollbars=no,resizable=no, menubar=no")
	newWin2.moveTo(tposx,tposy);
	newWin2.focus();
}

function winpopupflash(urlx,param,twidth,theight)
{
	var strurl= urlx + '?param1=' + param + '&param2=' +twidth+ '&param3=' + theight;
	var tposx= (screen.width- twidth)/2
	var tposy= (screen.height- theight)/2;

	var newWinflash=window.open(strurl,"windowflash","toolbar=no,width="+ twidth+",height="+ theight+ ",directories=no,status=no,scrollbars=yes,resizable=no, menubar=no")
	newWinflash.moveTo(tposx,tposy);
	newWinflash.focus();
}

function myOpenWindow(urlx,twidth,theight) {
    myWindowHandle = window.open(urlx,'myWindowName',"toolbar=no,width="+ twidth+",height="+ theight+ ",directories=no,status=no,scrollbars=yes,resizable=no, menubar=no");
}

//Begin function swap images
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//End function swap images

// Javascript hien thi Image dung kich co cua cua so Popup.
function openImage(ImageName) {
  PopUpWindow=window.open("","ImagePreview","menubar=0,status=0,scrollbars=1,directories=1,resizable=1,top=80,left=150");     
  	  // open the new window with just the title and status bar and name it 'PopUpWindow'
  PopUpWindow.document.writeln('<html>');
  PopUpWindow.document.writeln('<head>');
  PopUpWindow.document.writeln('<title>xsoft.com.vn</title>');   
  	  // put the name of the pic in the title bar
  PopUpWindow.document.writeln('</head>');

  if (navigator.appName == "Microsoft Internet Explorer") 
  	  // in IE resizeTo give the outer of the window
     PopUpWindow.document.writeln('<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" onLoad="window.resizeTo(document.images[0].width+50,document.images[0].height+ 100)">');   
		      // resize the window to match the picture

   else
       PopUpWindow.document.writeln('<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" onLoad="window.resizeTo(document.images[0].width,document.images[0].height+ 100)">');   
          // resize the window to match the picture
  PopUpWindow.document.writeln('<center><img src="'+ImageName+' " border="0">');
  PopUpWindow.document.writeln("</body></html>");

// load the image in the window
  PopUpWindow.document.writeln('</body>');
  PopUpWindow.document.writeln('</html>');
  PopUpWindow.document.close();
  PopUpWindow.focus();  // place the window in front
}
//popup chay Videoclips
function Avpopup(file_name,id,av_width,av_height)
{        
        window.open(file_name + "?param=" + id,'','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,width=' + av_width + ',height=' + av_height);
}

//Vote Submit
function Vote_submit(fForm)
{
	var i;
	var itemvalue;
	
	itemvalue="";
	for (i=0;i<fForm.VoteItem.length;i++)
	{	
		if (fForm.VoteItem[i].checked)
		{
			itemvalue+= fForm.VoteItem[i].value + " ";
			fForm.VoteItem[i].checked = false;
		}
	}
	window.open("/vote.asp?param=" + itemvalue + "&VoteId=" + fForm.VoteID.value, "Ket_Qua_Tham_Do", "toolbar=no,location=no,directories=no, status=no, menubar=no, scrollbars=yes,resizable=no,copyhistory=yes,width=400,height=300")
}
function Vote_view(fForm)
{
	window.open("/vote.asp?VoteId=" + fForm.VoteID.value, "Ket_Qua_Tham_Do", "toolbar=no,location=no,directories=no, status=no, menubar=no, scrollbars=yes,resizable=no,copyhistory=yes,width=400,height=300")
}
//doClick Chuong trinh giai tri
function doClick(id)
{	if (document.all(id,0).style.display == "none")
	{ document.all(id,0).style.display = ""; }
	else { document.all(id,0).style.display = "none";}
}

	function showAjax(strContent)
	{
		tipobj = document.createElement('DIV'); 
		tipobj.className = 'VietAdTooltip';
		tipobj.id = 'VietAdTooltip';		
		tipobj.style.display='block';
		tipobj.style.position = 'absolute';
		tipobj.innerHTML = strContent;
		document.body.appendChild(tipobj);
	}
	
	function hideAjax()
	{
		divObject = document.getElementById('VietAdTooltip')
		document.body.removeChild(divObject);
	}
