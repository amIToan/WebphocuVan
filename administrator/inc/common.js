// JavaScript Document
function winpopup(urlx,param,twidth,theight)
{
	var strurl= urlx + '?param=' + param;
	var tposx= (screen.width- twidth)/2
	var tposy= (screen.height- theight)/2;

	var newWin=window.open(strurl,"newWindow","toolbar=no,width="+ twidth+",height="+ theight+ ",directories=no,status=no,scrollbars=yes,resizable=no, menubar=no")
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

function IsNumeric(sText)
{
   //Kiem tra day co' phai la` 0 < sText <30.000
   
   var ValidChars = "0123456789";
   var IsNumber=true;
   var Char;
   if ((sText.length==0) || (sText.length>5))
   	   	return false;

   for (i = 0; i<sText.length; i++)
   {
      Char = sText.charAt(i); 
      if (ValidChars.indexOf(Char) == -1) 
         return false;
   }
   return true;
}

function CheckSubmit()
{
	if (document.fUpdate.AnswerFromWebAdmin.value == '')
	{
		alert('Đề nghị nhập nội dung trả lời!');
		document.ContactUs.AnswerFromWebAdmin.focus();
		return;
	}
	vWH = 160;
	vWW = 330;
	vWN = 'Discuss_Reply';
	winDef = 'status=no,resizable=no,scrollbars=no,toolbar=no,location=no,fullscreen=no,titlebar=yes,height='.concat(vWH).concat(',').concat('width=').concat(vWW).concat(',');
	winDef = winDef.concat('top=').concat((screen.height - vWH)/2).concat(',');
	winDef = winDef.concat('left=').concat((screen.width - vWW)/2);
	newwin = open('', vWN, winDef);

	document.fUpdate.action = 'test.html';
	document.fUpdate.target = vWN;
	document.fUpdate.submit();
}
