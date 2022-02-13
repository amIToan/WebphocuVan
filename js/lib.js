
//Start ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// Check Password Strenght

function checkPassword(strPassword)
{
	var bCheckNumbers = true;
	var bCheckUpperCase = true;
	var bCheckLowerCase = true;
	var bCheckPunctuation = true;
	var nPasswordLifetime = 365;
	// Reset combination count
	nCombinations = 0;
	
	// Check numbers
	if (bCheckNumbers)
	{
		strCheck = "0123456789";
		if (doesContain(strPassword, strCheck) > 0) nCombinations += strCheck.length; 
	}
	
	// Check upper case
	if (bCheckUpperCase)
	{
		strCheck = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
		if (doesContain(strPassword, strCheck) > 0) nCombinations += strCheck.length; 
	}
	
	// Check lower case
	if (bCheckLowerCase)
	{
		strCheck = "abcdefghijklmnopqrstuvwxyz";
		if (doesContain(strPassword, strCheck) > 0) nCombinations += strCheck.length;
	}
	
	// Check punctuation
	if (bCheckPunctuation)
	{
		strCheck = ";:-_=+\|//?^&!.@$£#*()%~<>{}[]";
		if (doesContain(strPassword, strCheck) > 0) nCombinations += strCheck.length; 
	}
	// Calculate
	// -- 500 tries per second => minutes 
    var nDays = ((Math.pow(nCombinations, strPassword.length) / 500) / 2) / 86400;
	// Number of days out of password lifetime setting
	var nPerc = nDays / nPasswordLifetime;
	
	return nPerc;
}
 
// Runs password through check and then updates GUI 
function runPassword(strPassword, strFieldID,arrText,arrColor) 
{
	
	var ctlBar = $("#"+strFieldID); 
	if (!ctlBar)return;
	
	// Check password
	if(strPassword!='')
	{
		nPerc = checkPassword(strPassword);
		
		// Set new width
		var nRound = Math.round(nPerc * 100);
		if (nRound < (strPassword.length * 5)) nRound += strPassword.length * 5; 
		if (nRound > 100) nRound = 100;
		//ctlBar.style.width = nRound + "%";
	 
	 // Color and text
		if (nRound > 95)
		{
			strText = arrText[0];
			strColor = arrColor[0];
		}
		else if (nRound > 65)
		{
			strText = arrText[1];
			strColor = arrColor[1];
		}
		else if (nRound > 30)
		{
			strText = arrText[2];
			strColor = arrColor[2];
		}
		else
		{
			strText = arrText[3];
			strColor = arrColor[3];
		}
		strBar='<div style="width: 65px;"><div style="font-size: 12px;color:'+strColor+'">&nbsp;' + strText + '</div><div style="font-size: 1px; height: 2px; width:'+nRound+'%; border: 1px solid white;background-color:'+strColor+'"></div></div>';
		
	}else strBar='&nbsp;*';
	ctlBar.html(strBar);
}
// Checks a string for a list of characters
function doesContain(strPassword, strCheck)
 {
    nCount = 0; 
	for (i = 0; i < strPassword.length; i++) 
	{
		if (strCheck.indexOf(strPassword.charAt(i)) > -1) 
		{ 
	        nCount++; 
		} 
	} 
 	
	return nCount; 
} 
//End ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

function RefreshLink()
{
	var url=window.location.href;
	var hash=window.location.hash;
	if(hash.length>2&&hash.indexOf("#")!=-1)
	{
		arrU=url.split("?");
		arrQ=arrU[1].split("#");
		strP=arrQ[0];
		strH=arrQ[1];
		var strUrl='';
		var listAdd='';
		var arrP_P=Array();
		var arrP_V=Array();
		arrP=strP.split("&");
		for(i=0;i<arrP.length;i++)
		{
			arrTemp1=arrP[i].split("=");
			arrP_P[i]=arrTemp1[0];
			arrP_V[i]=arrTemp1[1];
		}
		arrH=strH.split("&");
		for(j=0;j<arrH.length;j++)
		{
			arrTemp=arrH[j].split("=");
			bAdd=true;
			for(i=0;i<arrP.length;i++)
			{
				if(arrTemp[0]==arrP_P[i])
				{
					arrP_V[i]=arrTemp[1];
					bAdd=false;
				}
			}
			if(bAdd) listAdd+='&'+arrTemp[0]+"="+arrTemp[1];
		}
		for(i=0;i<arrP.length;i++) strUrl+='&'+arrP_P[i]+'='+arrP_V[i];
		strUrl='?'+strUrl.substring(1)+listAdd;
		window.location.href=strUrl;
	}
}
//RefreshLink();

function isEmail(_form,id,strAlert)
{
	var s=$.trim($('#'+_form+' input[name='+id+']').val());
	var i = 1;
	var sLength = s.length;
	var check=0;
	
	if (s=="") check+=1;
	if(s.indexOf(" ")>0) check+=1;
	if(s.indexOf("@")==-1) check+=1;
	if (s.indexOf(".")==-1) check+=1;
	if (s.indexOf("..")!=-1) check+=1;
	if (s.indexOf("@")!=s.lastIndexOf("@")) check+=1;
	if (s.lastIndexOf(".")==s.length-1) check+=1;
	var str="abcdefghikjlmnopqrstuvwxyz-@._1234567890";
	for(var j=0;j<s.length;j++)
	if(str.indexOf(s.charAt(j))==-1) check+=1;
	if(check)
	{
		alert(strAlert);
		$('#'+_form+' input[name='+id+']').focus();
		return false;
	}
	return true;
}


function GoUrl(url){window.location.href=url;}

function OpenViewImgs(form_,obj,folder)
{
	var strUrl="viewimgs.php?form="+form_+"&obj="+obj+"&fder="+folder;
	window.status="Open";
	window.open(strUrl,"View","scrollbars=yes,width=650,height=400");
}

function OpenFileManager(id)
{
	window.SetUrl=function(val)
	{
		val=val.replace(/[a-z]*:\/\/[^\/]*/,'');
		$("#"+id).val(val);
	};
	window.open('editor/plugins/kfm/?lang=vi','kfm','width=700,height=500');
}


function SelectChange(obj)
{if(obj.options[obj.selectedIndex].value != 0) window.location.href=obj.options[obj.selectedIndex].value;}

function CheckDel(){return QuestionDel('')}
function QuestionDel(str)
{
	if(str=='') str='Bạn chắc chắn muốn xoá không?';
	return (confirm(str))
}

function OpenWin(strUrl,Name,Boder){window.open(strUrl,Name,Boder)}
function changeto(obj,strClass){
	if(strClass!=""){
		obj.className = strClass;
	}
	obj.style.cursor = 'hand';
}

function ObjectExist(id){return ($("#"+id)==undefined)?false:true;}

function CheckBoxAll(_form,chkBox,type)
{
	/*
	type=0: Check All
	type=1: Uncheck All
	*/
	var els = $("form :checkbox");
	for(i=0; i<els.length; i++)
	{ 
		if(els[i].name.substr(0,chkBox.length)==chkBox&&!els[i].disabled)
		{
			if(type==0) els[i].checked=true;
			else els[i].checked=false;
		}
	}
}

function CheckDisableAll(_form,chkBox,type)
{
	/*
	type=0: Disbale All
	type=1: UnDisable All
	*/
	var els = $("form :checkbox");
	for(i=0; i<els.length; i++)
	{ 
		if(els[i].name.substr(0,chkBox.length)==chkBox)
		{
			if(type==0) els[i].disabled=true;
			else els[i].disabled=false;
		}
	}
}

function GetListCheckboxValue(_form,chkBox)
{
	var list='';
	var els = $("form :checkbox");
	j=0;
	for(i=0; i<els.length; i++)
	{ 
		if(els[i].name.substr(0,chkBox.length)==chkBox&&els[i].checked)
		{
			if(j==0) list+=els[i].value; else list+=","+els[i].value;
			j++;
		}
	}
	return list;
}

// Start Listbox Multi Sort
function MoveUpDown(name,w,s)
{
	var sel=document.getElementById(name);
	var idx=sel.selectedIndex;
	if(idx==-1) alert(s);
	else
	{
		var opt=sel.options[idx];
		if(w=='up')
		{
			var prev=opt.previousSibling;
			while(prev&&prev.nodeType!=1)prev=prev.previousSibling;
			prev?sel.insertBefore(opt,prev):sel.appendChild(opt)
		}
		else
		{
			var next=opt.nextSibling;
			while(next&&next.nodeType!=1)next=next.nextSibling;
			if(!next)
				sel.insertBefore(opt,sel.options[0])
			else
			{
				var nextnext=next.nextSibling;
				while(next&&next.nodeType!=1) next=next.nextSibling;
				nextnext?sel.insertBefore(opt,nextnext):sel.appendChild(opt);
			}
		}
	}
}

function PostListValue(name,idto) 
{
	var box=document.getElementById(name);
	var va='';
	for(var i=0; i<box.length; i++) 
	{
		va+=box.options[i].value+',';
	}
	document.getElementById(idto).value=va;
}  
// End Listbox Multi Sort

// NDK Loading

var timershow=false;
var curx=-200;
var cury=350;
var win_w=window.innerWidth ? window.innerWidth : document.body.offsetWidth;
var win_h=window.screenHeight? window.innerHeight: document.body.offsetHeight;
var mid_w=win_w/2-100;
var mid_h=win_h/2+20;

function show_Loading() {
	obj = $("#LoadingDiv");
	//alert(obj.left);
	obj.css("left",mid_w + "px");
	obj.css("top",mid_h+ "px");
}

function hide_Loading() {
	obj = $("#LoadingDiv");
	obj.css("left",curx + "px");
	obj.css("top",cury+ "px");
}
// End

function AddHref(param)
{
	window.location.hash=param;
}

function CheckField(_form,listfiled,strAlert)
{
	arrTmp=listfiled.split(',');
	for(i=0;i<arrTmp.length;i=i+2)
	{
		if($('#'+_form+' [name='+arrTmp[i]+']').val()==arrTmp[i+1])
		{
			alert(strAlert);
			$('#'+_form+' [name='+arrTmp[i]+']').focus();
			return false;
		}
	}
	return true;
}

function CheckFieldIfOne(_form,listfiled,strAlert)
{
	if(listfiled!="")
	{
		arrTmp=listfiled.split(',');
		check=false;
		for(i=0;i<arrTmp.length;i=i+2)
		{
			if($('#'+_form+' [name='+arrTmp[i]+']').val()!=arrTmp[i+1])
			{
				check=true;
				break;
			}
		}
		if(!check)
		{
			alert(strAlert);
			$('#'+_form+' [name='+arrTmp[0]+']').focus();
			return false;
		}
		return true;
	}
	else
	{
		alert('List check is null');
		return false;
	}
}

function CheckFieldNumber(_form,listfiled,strAlert)
{
	arrTmp=listfiled.split(',');
	for(i=0;i<arrTmp.length;i++)
	{
		if(isNaN($('#'+_form+' [name='+arrTmp[i]+']').val()))
		{
			alert(strAlert);
			$('#'+_form+' [name='+arrTmp[i]+']').focus();
			return false;
		}
	}
	return true;
}

function CheckUrl(strUrl)
{
	var RegexUrl = /(ftp|http|https):\/\/(\w+:{0,1}\w*@)?(\S+)(:[0-9]+)?(\/|\/([\w#!:.?+=&%@!\-\/]))?/
	return RegexUrl.test(strUrl);
}


function ShowTabClick(path)
{
	var ctl=$('#ShowTabLeft');
	var ConfigTab;
	if(ctl.attr('class')=='Show')
	{
		ctl.attr({'class':'Hide'})
		$('#ImageTabShow').attr("src",path+"images/bar_open.gif");
		ConfigTab=1;
	}else
	{
		ctl.attr({'class':'Show'})
		$('#ImageTabShow').attr("src",path+"images/bar_close.gif");
		ConfigTab=0;
	}
	$.ajax({type: "GET",url:"admin_ajax.php",data:'modul=main&sub=ajax&method=tab&cfg='+ConfigTab});
}

function ClearField(id,_form)
{
	if(_form=='') $("#"+id).val('');
	else $('#'+_form+' input[name='+id+']').val('');
}
function ShowHideById(ShowId,HideId)
{
	$("#"+ShowId).slideDown();
	$("#"+HideId).slideUp();
}

function ShowContentById(id,ctlID)
{
	var IdHide=$("#"+ctlID).val();
	$("#"+ctlID).val(id);
	if(IdHide!=id) 
	{
		$("#"+id).slideDown();
		if(IdHide!='') $("#"+IdHide).slideUp();
	}
	else ($("#"+id).css("display")!='none')?$("#"+id).slideUp():$("#"+id).slideDown();
}
function LoadAjaxPage(method,url,param,id,bRedirect,ulrRedirect)
{
	var RedirectUrl=document.location.href;
	param=param.replace("?","");
	if(bRedirect==true&&ulrRedirect!="") RedirectUrl=ulrRedirect;
	show_Loading();
	//alert(url);
	$.ajax(
	{
		type: method, 
		url: url,
		data: param, 
		success: function(transport)
		{
			textValue=$.trim(transport);
			//alert(textValue);
			hide_Loading();
			if(textValue=="OK"&&bRedirect==true)
			{
				document.location.href=RedirectUrl;
			}else
			{
				if(textValue.substring(0,6)=="ALERT!")
				{
					alert(textValue.substring(6));
				}else
				{
					if(textValue.substring(0,3)=="OK!"&&bRedirect==true)
					{
						alert(textValue.substring(3));
						document.location.href=RedirectUrl;
					}else
					{
						$("#"+id).html(textValue);
					}
				}
			}
		}
	});
}

function AjaxLoad(url,param,id)
{
	LoadAjaxPage('GET',url,param,id,false,'');
}
function AjaxLoad1(url,param,id,bRedirect)
{
	LoadAjaxPage('GET',url,param,id,bRedirect,'');
}
function AjaxLoad2(url,param,id,bRedirect,ulrRedirect)
{
	LoadAjaxPage('GET',url,param,id,bRedirect,ulrRedirect);
}

/*function GetAjaxDataPost(_form,param)
{
	var els =$(_form).elements; 
	query='rand='+parseInt(Math.random()*99999999)+'&'+param;
	for(i=0; i<els.length; i++)
	{
		if(els[i].type=="checkbox"&&(!(els[i].checked)||els[i].disabled))
			query+='&'+els[i].name+'=';
		else
		{
			if(els[i].type!="radio")
			{
				value=els[i].value;
				value=value.replace(/#/g,"[0023;]");
				value=value.replace(/&/g,"[0026;]");
				value=value.replace(/\?/g,"[003F;]");
				//value=value.replace("\\","[+005C;]");
				query+='&'+els[i].name+'='+encodeURI(value);
			}else
			if(els[i].type=="radio"&&els[i].checked&&!els[i].disabled)
				query+='&'+els[i].name+'='+encodeURI(els[i].value);
		}
	}
	return query;
}
*/function AjaxPost(_form,url,param,id,bRedirect,ulrRedirect)
{
	var query =$.param($("#"+_form).serializeArray())+"&"+param; 
	//alert(query);
	LoadAjaxPage('POST',url,query,id,bRedirect,ulrRedirect);
}
function AjaxPostWithLoad(_form,url,param,id,url1,param1,id1)
{
	show_Loading();
	var query =GetAjaxDataPost(_form,param); 
	var Load=false;
	var myAjax = new Ajax.Request(
	url, 
	{
		method: 'post', 
		parameters: query, 
		onComplete: function(transport)
		{
			textValue=trim(transport.responseText);
			hide_Loading();
			if(textValue=="OK")
			{
				$(id).innerHTML='';
				Load=true;
			}else
			{
				if(textValue.substring(0,6)=="ALERT!")
				{
					alert(textValue.substring(6));
				}else
				{
					if(textValue.substring(0,3)=="OK!")
					{
						Load=true;
						$(id).innerHTML='';
						alert(textValue.substring(3));
					}else
					{
						$(id).innerHTML=textValue;
					}
				}
			}
			if(Load)
			{
				$(_form).disable();
				AjaxLoad(url1,param1,id1);
			}
		}
	});
}

function RefreshCaptcha(id)
{
	var src=$("#"+id).attr('src');
	$("#"+id).attr('src',src+'?rnd='+Math.random());
}

function ChangeNumberPerPage(ctl,from,url,param,id,drect,urldrect)
{
	param=param+"&"+ctl.name+"="+ctl.value;
	AjaxPost(from,url,param,id,drect,urldrect);
}
function LoadCheckValueExist(method,url,param,id)
{
	show_Loading();
	$.ajax(
	{
		type: method, 
		url: url,
		data: param, 
		success: function(transport)
		{
			textValue=$.trim(transport);
			hide_Loading();
			if(textValue=="TRUE") text='<img src="images/true.gif" width="16px" height="16px">';
			else if(textValue=="FALSE") text='<img src="images/false.gif" width="16px" height="16px">';
			$("#"+id).html(text);
		}
	});
}

function CheckValueExist(form,field,url,param,id,defcheck,defshow)
{
	valueCheck=$("#"+form+" input[name="+field+"]").val();
	if($.trim(valueCheck)!=defcheck)
	{
		param+="&"+field+"="+encodeURI(valueCheck);
		LoadCheckValueExist('get',url,param,id);
	}else $("#"+id).html(defshow);
}

function ajaxSelect2Select(_form,selectbox1,selectbox2,url,param,optdef,iddef,titledef)
{
	var selbox1=$("#"+_form+" select[name="+selectbox1+"]");
	var sel=selbox1.val();
	$("#"+_form+" select[name="+selectbox2+"]").html('<option value="">Loading...</option>');
	$("#"+_form+" select[name="+selectbox2+"]").attr({"disabled":"disabled"});
	$.ajax(
	{
		type: 'GET', 
		url: url,
		data:  param+'&'+selectbox1+'='+sel,
		dataType:"json",
		success: function(data)
		{
			//alert(data);
			var options = (optdef==true)?'<option value="'+iddef+'">'+titledef+'</option>':'';
			$.each(data.lists, function(i,opt)
			{
				options += '<option value="' + opt.id + '">' + opt.title + '</option>';
			});
			$("#"+_form+" select[name="+selectbox2+"]").html(options);
			$("#"+_form+" select[name="+selectbox2+"]").attr({"disabled":""});
		}
	});
}
/*function removeAllOptions(selectbox)
{
	var selbox=$("#"+selectbox);
    for(i=selbox.options.length-1;i>=0;i--){selbox.remove(i);}
}*/

function CompareString(str1,str2,lowcase)
{
	if(lowcase=true)
	{
		str1=str1.toLowerCase();
		str2=str2.toLowerCase();
	}
	if(str1==str2) return true;
	else return false;
}

function CheckValue2Field(form,field1,field2,id,defcheck,defshow)
{
	var val1=$("#"+form+" input[name="+field1+"]").val();
	var val2=$("#"+form+" input[name="+field2+"]").val();
	if(!(val1==defcheck&&val2==defcheck))
	{
		if(CompareString(val1,val2,false))
			text='<img src="images/true.gif" width="16px" height="16px">';
		else text='<img src="images/false.gif" width="16px" height="16px">';
		$("#"+id).html(text);
	}else $("#"+id).html(defshow);
}

function CheckFieldValue(form,field,idshow,defcheck,defshow)
{
	if($("#"+form+" input[name="+field+"]").val()==defcheck) $("#"+idshow).html(defshow);
	else $("#"+idshow).html('<img src="images/true.gif" width="16px" height="16px">');
}

//Thoi tiet
function AWeather(_form,IdPlugins,ImgPath)
{
	var vID=$('#'+_form+'_'+IdPlugins+' select[name=cboWeather_'+IdPlugins+']').val();
	if (vID==1){vFile="Sonla.xml";}		
	else if (vID==2){vFile="Viettri.xml";}
	else if(vID==3){vFile="Haiphong.xml";}
	else if(vID==4){vFile="Hanoi.xml";}
	else if(vID==5){vFile="Vinh.xml";}
	else if(vID==6){vFile="Danang.xml";	}
	else if(vID==7){vFile="Nhatrang.xml";}
	else if(vID==8){vFile="Pleicu.xml";}
	else{vFile="HCM.xml";}
	$.ajax(
	{
		type: 'GET', 
		url: 'ajax.php',
		data: 'modul=common&sub=ajax&method=weather&file='+vFile, 
		success: function(req)
		{
			var vAdImg; var vAdImg1; var vAdImg2; var vAdImg3; var vAdImg4; var vAdImg5; var vWeather;
			var arr=req.split(' - ');
			vAdImg = arr[0];
			vAdImg1 = arr[1];
			vAdImg2 = arr[2];
			vAdImg3 = arr[3];
			vAdImg4 = arr[4];
			vAdImg5 = arr[5];
			vWeather = arr[6];
			AdWeather(_form,IdPlugins,ImgPath,vAdImg,vAdImg1,vAdImg2,vAdImg3,vAdImg4,vAdImg5,vWeather);
		}
	});
}

function AdWeather(_form,IdPlugins,ImgPath,vImg,vImg1,vImg2,vImg3,vImg4,vImg5,vWeather){
	var AdDo;
	AdDo = "<img src='"+ImgPath+"weather/" + vImg1 + "' border='0'>";
	if (vImg2 != '') AdDo += "<img src='"+ImgPath+"weather/" + vImg2 + "' border='0'>";
	if (vImg3 != '') AdDo += "<img src='"+ImgPath+"weather/" + vImg3 + "' border='0'>";
	if (vImg4 != '') AdDo += "<img src='"+ImgPath+"weather/" + vImg4 + "' border='0'>";
	if (vImg5 != '') AdDo += "<img src='"+ImgPath+"weather/" + vImg5 + "' border='0'>";
	AdDo +="<img src='"+ImgPath+"weather/c.gif' border='0'>";
	$("#TextWeather_"+IdPlugins).html(vWeather);
	$("#ImgWeather_"+IdPlugins).html("<img src='http://www.vnexpress.net/Images/Weather/" + vImg + "' border='0'>");
	$("#ImgC_"+IdPlugins).html(AdDo);
}

//Gia Vang
function ShowGoldPrice(IdPlugins)
{
	$("#SbjBuy_"+IdPlugins).html(vGoldSbjBuy);
	$("#SbjSell_"+IdPlugins).html(vGoldSbjSell);
	$("#SjcBuy_"+IdPlugins).html(vGoldSjcBuy);
	$("#SjcSell_"+IdPlugins).html(vGoldSjcSell);
}

//Ty gia
function ShowForexRate(IdPlugins)
{
	var sHTML = '';
	sHTML = sHTML.concat('<table width="99%" border="1" cellspacing="0" cellpadding="3" style="background-color:#FFFFFF;border:1px solid #999999;border-collapse:collapse;" bordercolor="#999999" align="center">');
	for(var i=0;i<vForexs.length;i++){
		if(vForexs[i]!="")
		{
			sHTML = sHTML.concat('	<tr>');
			sHTML = sHTML.concat('		<td class="tdGold">').concat(vForexs[i]).concat('</td>');
			sHTML = sHTML.concat('		<td class="tdGold">').concat(vCosts[i]).concat('</td>');
			sHTML = sHTML.concat('	</tr>');
		}
	}
	sHTML = sHTML.concat('</table>');
	$("#eForex_"+IdPlugins).html(sHTML);
}

//Binh chon
function CheckThisVote(_form,ctl)
{
	MaxChoose=parseInt($('#'+_form+' input[name=MaxChoose]').val());
	Choose=parseInt($('#'+_form+' input[name=Choose]').val());
	ListVote=$('#'+_form+' input[name=ListVote]').val();
	
	value=ctl.value;
	if(ctl.checked)
	{
		if(Choose<MaxChoose)
		{
			
			ListVote+=value+',';
			Choose++;
			$('#'+_form+' input[name=ListVote]').val(ListVote);
			$('#'+_form+' input[name=Choose]').val(Choose);
		}else ctl.checked=false;
	}else
	{
		ListVote=ListVote.replace(','+value+',',',');
		Choose--;
		$('#'+_form+' input[name=ListVote]').val(ListVote);
		$('#'+_form+' input[name=Choose]').val(Choose);
	}
	
}
function CheckPoll(_form,strAlert)
{
	Choose=parseInt($('#'+_form+' input[name=Choose]').val());
	ListVote=$('#'+_form+' input[name=ListVote]').val();
	if(Choose==0||ListVote==',')
	{
		alert(strAlert);
		return false;
	}
}

function SubmitForm(_form){$("#"+_form).submit();}
function DefaultValue(ctl,value)
{
	if($.trim(ctl.value)==''||$.trim(ctl.value)==value) ctl.value=value
}
function ClearValue(ctl,value)
{
	if($.trim(ctl.value)==value) ctl.value=''
}

function ChangePathOther(ctl,id,other)
{
	$("#"+id).html("<span class=\"h7\">"+other+"</span> "+$("#"+ctl).html());
}

var menu=function(){
	var t=15,z=50,s=6,a;
	function dd(n){this.n=n; this.h=[]; this.c=[]}
	dd.prototype.init=function(p,c){
		a=c; var w=document.getElementById(p), s=w.getElementsByTagName('ul'), l=s.length, i=0;
		for(i;i<l;i++){
			var h=s[i].parentNode; this.h[i]=h; this.c[i]=s[i];
			h.onmouseover=new Function(this.n+'.st('+i+',true)');
			h.onmouseout=new Function(this.n+'.st('+i+')');
		}
	}
	dd.prototype.st=function(x,f){
		var c=this.c[x], h=this.h[x], p=h.getElementsByTagName('a')[0];
		clearInterval(c.t); c.style.overflow='hidden';
		if(f){
			p.className+=' '+a;
			if(!c.mh){c.style.display='block'; c.style.height=''; c.mh=c.offsetHeight; c.style.height=0}
			if(c.mh==c.offsetHeight){c.style.overflow='visible'}
			else{c.style.zIndex=z; z++; c.t=setInterval(function(){sl(c,1)},t)}
		}else{p.className=p.className.replace(a,''); c.t=setInterval(function(){sl(c,-1)},t)}
	}
	function sl(c,f){
		var h=c.offsetHeight;
		if((h<=0&&f!=1)||(h>=c.mh&&f==1)){
			if(f==1){c.style.filter=''; c.style.opacity=1; c.style.overflow='visible'}
			clearInterval(c.t); return
		}
		var d=(f==1)?Math.ceil((c.mh-h)/s):Math.ceil(h/s), o=h/c.mh;
		c.style.opacity=o; c.style.filter='alpha(opacity='+(o*100)+')';
		c.style.height=h+(d*f)+'px'
	}
	return{dd:dd}
}();


function addCompareList(productId) {
    var checked = document.getElementById("compareItem_"+productId).checked;
    var currentList = $("#productCompareList").val();
    var currentNumItem = currentList.split(";").length - 2;
    var productImageUrl = $("#productImg_"+productId).attr("src");
	//alert(productImageUrl);
    var iconHtml = "<img src=\""+ productImageUrl + "\" width=20 height=20>";
    if (checked) {
        if (currentNumItem > 2) {
            //Cho phep so sánh tối đa 3 sản phẩm
            document.getElementById("compareItem_"+productId).checked = "";
            alert ("Bạn chỉ có thể so sánh tối đa 3 sản phẩm");
        }
        else
        {
            $("#productCompareList").val(currentList + productId + ";");
            //Thêm tiếp ảnh
            nextContainer = currentNumItem + 1;
            $("#compareItemContain_"+ nextContainer).html(iconHtml);
            $("#productItemContain_"+ nextContainer).val(productId);
        }
    }
    else
    {
        $("#productCompareList").val(currentList.replace(";" + productId + ";",";"));
        //Xóa bỏ ảnh khỏi compare
        var containId = 0;
        for(var i=1; i<=currentNumItem; i++) {
            var productContainer = $("#productItemContain_"+ i).val();
            if (productContainer != ""){
                if (productId == productContainer) {
                    $("#compareItemContain_"+ i).html("");
                    $("#productItemContain_"+ i).val("");
                    containId = i;
                    break;
                }
            }
        }
        //Sắp xếp lại ảnh
        if(containId != 0){
            arrangeImageContainer(containId,3);
        }
    }
}

//Sắp xếp lại ảnh trong các container đảm bảo các ảnh từ 1 -> n
//startId là ID của container có ảnh bị remove, do đó cần được bù bằng ảnh khác nằm ở container lớn hơn gần nhất nếu có
//rồi lặp lại quá trình này cho container vừa bị lấy ảnh

function arrangeImageContainer(startId,maxNum){
    var nextStartId = 0;
    for (var i=startId+1; i<=maxNum; i++){
        var contentHtml = $("#compareItemContain_"+ i).html();
        var productContainer = $("#productItemContain_"+ i).val();
        if (contentHtml.length > 2)
        {
            $("#compareItemContain_"+ startId).html(contentHtml);
            $("#productItemContain_"+ startId).val(productContainer);
            $("#compareItemContain_"+ i).html("");
            $("#productItemContain_"+ i).val("");
            nextStartId = i;
            arrangeImageContainer(nextStartId,maxNum);
            break;
        }
    }
}

function ShowTooltip(id){showtip($('#'+id).html());}

function ShowBlockOverlay(id,width,height)
{
	$.blockUI({ message: $('#'+id), 
		css: { 
			top:  ($(window).height() - height) /2 + 'px', 
			left: ($(window).width() - width) /2 + 'px', 
			width: width+'px',
			height: height +'px',
			cursor: 'default'
		},
		overlayCSS: { cursor: 'default' }
	 }); 
	 $('.blockOverlay').attr('title','Click to unblock').click($.unblockUI);
}