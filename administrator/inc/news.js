	function checkFormData(sVar)
	{




	    if ((document.fInsert.CatId_DependRole.value == 0) && (document.fInsert.CatId_DependRole.value == "") && (sVar == 1))
		{
			alert('Chưa chọn \"Chuyên mục\"');
			document.fInsert.CatId_DependRole.focus();
			return false;
		}
		//Title
		if (document.fInsert.Title.value == "")
		{
			alert('Chưa nhập \"Tiêu đề\"');
			document.fInsert.Title.focus();
			return false;
		}
		//Description


		if ((document.fInsert.StatusId.value == "0") && (sVar==1))
		{
			alert('Lựa chọn \"Trạng thái đăng tin\"');
			document.fInsert.StatusId.focus();
			return false;
		}

		if (document.getElementById("attach_product").checked== false) {
		    alert('Lựa chọn \"Xác nhận đăng tin.\"');
		    document.fInsert.attach_product.focus();
		    return false;
		}
		return true;
		
	}
	
	function preview()
	{
		if (!checkFormData(0)) 
			return;
		
		vWH = 450;
		vWW = 720;
		vWN = 'Preview';
		winDef = 'status=no,resizable=yes,scrollbars=yes,toolbar=no,location=no,fullscreen=no,titlebar=yes,height='.concat(vWH).concat(',').concat('width=').concat(vWW).concat(',');
		winDef = winDef.concat('top=').concat((screen.height - vWH)/2).concat(',');
		winDef = winDef.concat('left=').concat((screen.width - vWW)/2);
		newwin = open('', vWN, winDef);

		document.fInsert.action = 'news_preview.asp';
		document.fInsert.target = vWN;
		document.fInsert.submit();
	}
	
	function SendToOneCat()
	{
		if (!checkFormData(1)) 
			return;
		document.fInsert.action="news_up.asp";
		document.fInsert.target = "_self";
		document.fInsert.submit();
	}
	
	function SendToMultiCat()
	{
		if (!checkFormData(0)) 
			return;
		
		vWH = 340;
		vWW = 350;
		vWN = 'CheckCategory';
		winDef = 'status=no,resizable=yes,scrollbars=yes,toolbar=no,location=no,fullscreen=no,titlebar=yes,height='.concat(vWH).concat(',').concat('width=').concat(vWW).concat(',');
		winDef = winDef.concat('top=').concat((screen.height - vWH)/2).concat(',');
		winDef = winDef.concat('left=').concat((screen.width - vWW)/2);
		newwin = open('', vWN, winDef);
		
		document.fInsert.action = 'news_choosemulticat.asp';
		document.fInsert.target = vWN;
		document.fInsert.submit();
	}
	
	function Edit_preview()
	{
		if (!checkFormData(0)) 
			return;
		
		vWH = 450;
		vWW = 720;
		vWN = 'Preview';
		winDef = 'status=no,resizable=yes,scrollbars=yes,toolbar=no,location=no,fullscreen=no,titlebar=yes,height='.concat(vWH).concat(',').concat('width=').concat(vWW).concat(',');
		winDef = winDef.concat('top=').concat((screen.height - vWH)/2).concat(',');
		winDef = winDef.concat('left=').concat((screen.width - vWW)/2);
		newwin = open('', vWN, winDef);

		document.fInsert.action = 'news_edit_preview.asp';
		document.fInsert.target = vWN;
		document.fInsert.submit();
	}
	
	function Edit_SendToOneCat()
	{
		if (!checkFormData(1)) 
			return;
		document.fInsert.action="news_update.asp";
		document.fInsert.target = "_self";
		document.fInsert.submit();
	}
	
	function Edit_SendToMultiCat_Edit()
	{
		if (!checkFormData(0)) 
			return;
		
		vWH = 340;
		vWW = 350;
		vWN = 'CheckCategory';
		winDef = 'status=no,resizable=yes,scrollbars=yes,toolbar=no,location=no,fullscreen=no,titlebar=yes,height='.concat(vWH).concat(',').concat('width=').concat(vWW).concat(',');
		winDef = winDef.concat('top=').concat((screen.height - vWH)/2).concat(',');
		winDef = winDef.concat('left=').concat((screen.width - vWW)/2);
		newwin = open('', vWN, winDef);
		
		document.fInsert.action = 'news_choosemulticat.asp?var=update';
		document.fInsert.target = vWN;
		document.fInsert.submit();
	}
	
	function checkInteger(x)
	{
		if ((x==0)||(x==1)||(x==2)
		||	(x==3)||(x==4)||(x==5)
		||	(x==6)||(x==7)||(x==8)
		||	(x==9))
		{
			return 1 
		}
		else
			return 0;		
	}
	function DisMoney(strMoney)
	{
		var lg = strMoney.length;
		var tempStr='';
		var strTemp=strMoney;
		var x = strTemp.charAt(lg-1);
			
		if (checkInteger(x) == 0)
		{	
			strTemp = strTemp.substring(0,lg-1);				
			return strTemp;
		}
		for(k=1;k<=5;k++)
		{
			strMoney = strMoney.replace(",","")
		}
		lg = strMoney.length;
		du =	lg % 3
		iSo =(lg-du)/3
		sTien=''
		k = 0
		for(i=0;i<iSo;i++)
		{
			strTemp = strMoney.substring(lg -k -3,lg - k)
			sTien	=	',' + strTemp + sTien
			k= k+3
			}
		sTien = strMoney.substring(0,du) + sTien
		if ((sTien.charAt(0)==',')||(sTien.charAt(0)=='0'))
		{
			sTien=sTien.substring(1,sTien.length);
		}			
		return sTien;
	}
	
	function Edit_SendToMultiCat_Edit()
	{
		if (!checkFormData(0)) 
			return;
		
		vWH = 340;
		vWW = 350;
		vWN = 'CheckCategory';
		winDef = 'status=no,resizable=yes,scrollbars=yes,toolbar=no,location=no,fullscreen=no,titlebar=yes,height='.concat(vWH).concat(',').concat('width=').concat(vWW).concat(',');
		winDef = winDef.concat('top=').concat((screen.height - vWH)/2).concat(',');
		winDef = winDef.concat('left=').concat((screen.width - vWW)/2);
		newwin = open('', vWN, winDef);
		
		document.fInsert.action = 'news_choosemulticat.asp?var=update';
		document.fInsert.target = vWN;
		document.fInsert.submit();
	}
	
	var zoomfactor=0.05 
function zoomhelper()
{
	if (parseInt(whatcache.style.width)>10&&parseInt(whatcache.style.height)>10)
	{
		whatcache.style.width = parseInt(whatcache.style.width)
			+  parseInt(whatcache.style.width)*zoomfactor*prefix
		whatcache.style.height=parseInt(whatcache.style.height)
				+ parseInt(whatcache.style.height)*zoomfactor*prefix
	}
}
function zoom(originalW, originalH, what, state)
{
	if (!document.all&&!document.getElementById)
		return
	whatcache=eval("document.images."+what)
	prefix=(state=="in")? 1 : -1
	if (whatcache.style.width==""||state=="restore")
	{
		whatcache.style.width=originalW
		whatcache.style.height=originalH
		if (state=="restore")
			return
	}
	else
	{
		zoomhelper()
	}
	beginzoom=setInterval("zoomhelper()",100)
}
function clearzoom()
{
	if (window.beginzoom)
		clearInterval(beginzoom)
}
