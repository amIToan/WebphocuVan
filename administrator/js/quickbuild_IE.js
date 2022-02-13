/*** Freeware Open Source writen by ngoCanh 2002-05                  */
/*** Original by Vietdev  http://vietdev.sourceforge.net             */
/*** Release 2004-03-15  R9.0                                        */
/*** GPL - Copyright protected                                       */
/*********************************************************************/

function mouseOver(eButton)
	{
	eButton.className = "ButtonMouseOver";}
function mouseOut(eButton)
	{
	eButton.className = "ButtonNormal";	}
function mouseDown(eButton)
	{
	if (event.button == 1){
	eButton.className = "ButtonPressed";}	}
function mouseUp(eButton)
	{
	if (event.button == 1){
	eButton.className = "ButtonNormal";	eButton = null;} 
	}

function iEditor(idF)
{
  var obj= document.frames[idF]
  obj.document.designMode="On"

  obj.document.attachEvent("onmousedown", function(){ TXTOBJ=null; fID=idF;})
  obj.document.attachEvent("onkeydown", FKDown)

  var arr= idF.split("VDevID");
  var val= document.forms[arr[0]][arr[1]].value

   val= val.replace(/\r/g,"");
   val= val.replace(/\n</g,"<");
   
   var reg= /<pre>/i ;
   if( reg.test(val) )
	 { val= val.replace(/\n/g, "&#13;"); val= val.replace(/\t/g, "     "); }

   val= val.replace(/\n/g, "<br>");
   val= val.replace(/\t/g, "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;");

   val= val.replace(/\\/g, "&#92;");
   val= val.replace(/\'/g, "&#39;");

   if(val && val.indexOf('ViEtDeVdIvId')>=0) val= initDefaultOptions1(val,idF)
   else initDefaultOptions0(idF)

   setTimeout("document.frames['"+idF+"'].document.body.innerHTML='"+val+"'",200)


   TXTOBJ= null
   format[idF]='HTML'
   viewm[idF]=1;

   obj.focus();
}

function changetoIframeEditor(el)
{
   if( navigator.platform!="Win32" ) return null;

   var wi= '', hi= '';
   if(el.style.height) hi= el.style.height
   else if(el.rows) hi= (14*el.getAttribute('rows')+28)
   if(el.style.width) wi= el.style.width
   else if(el.cols) wi= (6*el.getAttribute('cols') +25)
   
   var parent= el.parentNode   

   
   while(parent.nodeName != 'FORM') parent= parent.parentNode
   var oform= parent
   var fidx=0; while(document.forms[fidx] != oform) fidx++ ; // form index


   var val='';

   if(el.nodeName=='TEXTAREA' || el.nodeName=='INPUT')
	 { fID= fidx+'VDevID'+el.getAttribute('name'); val= el.value }
   else fID= fidx+'VDevID'+el.getAttribute('id')


   createEditor(el,fID,wi,hi);

   setTimeout("iEditor('"+fID+"')",200); 
   return true;
  
}


//////////////////////////////
// for text mode
function doFormat(arr,caret)
{
  var wrd=TXTOBJ.curword.text

  var cmd = new Array();
  cmd = arr.split(',')

  if(!cmd[0]) return 
  if(cmd[0]=='SelectAll') { TXTOBJ.focus(); TXTOBJ.select(); return }
  if(cmd[0]=='Cut') { caret.execCommand("Cut"); return }
  if(cmd[0]=='Copy') { caret.execCommand("Copy"); return }
  if(cmd[0]=='Paste') { caret.execCommand("Paste"); return }

  TXTOBJ.curword=caret.duplicate();
  TXTOBJ.curword.text= cmd[0]+wrd+cmd[1]
}

// init all found TEXTAREA in document
function changeAllTextareaToEditors()
{
  var i=0;
  while(document.getElementsByTagName('textarea')[i])
   { 
    if(!changetoIframeEditor(document.getElementsByTagName('textarea')[i])) break;
	if(++i>0 && !document.getElementsByTagName('textarea')[i] ) i=0;
   }
}



// init all found IFRAME in document to Editable
function changeAllIframeToEditors()
{
  var i=0;
  while(document.getElementsByTagName('iframe')[i])
  { 
	if(!changetoIframeEditor(document.getElementsByTagName('iframe')[i])) break;
	i++
  }

}



// init some IFRAMEs
// e.g. changeIframeToEditor('id1','id2',...); // id1= id of frame
function changeIframeToEditor()
{
  for(var j=0;j<arguments.length;j++)
   {
     var i=0;
	 while(document.getElementsByTagName('iframe')[i])
	  { 
		if(document.getElementsByTagName('iframe')[i].id == arguments[j])
		  {	changetoIframeEditor(document.getElementsByTagName('iframe')[i]); break; }
	    i++
	  }
   }
}




/////////////////////////////////////////////////////////////////
function controlRows(fid)
{
	var str ="<TR valign=middle EVENT>\
<TD bgcolor='#dddddd'>\
<SELECT name='QBCNTRL0' title='TextStyle' onchange='setTextStyle(this.value)'>\
<OPTION value=''>"+ M_DEFTSTYLE +
"<OPTION value='" + M_TSTYLE1 + "'>" + M_TSTYLE1T + 
"<OPTION value='" + M_TSTYLE2 + "'>" + M_TSTYLE2T + 
"<OPTION value='" + M_TSTYLE3 + "'>" + M_TSTYLE3T +
"<OPTION value='" + M_TSTYLE4 + "'>" + M_TSTYLE4T +
"<OPTION value='" + M_TSTYLE5 + "'>" + M_TSTYLE5T +
"</SELECT>\
<SELECT name='QBCNTRL1' title='FontName' onchange='doFormatF(\"FontName,\"+this.value)'>\
<OPTION value=''>Font chữ" +
"<OPTION value='Arial' style='font-family:Arial'>Arial" +
"<OPTION value='Times New Roman' style='font-family:Times New Roman'>Times New Roman" +
"<OPTION value='Courier New' style='font-family:Courier New'>Courier New" +
"<OPTION value='Georgia' style='font-family:Georgia'>Georgia" +
"<OPTION value='Verdana' style='font-family:Verdana'>Verdana" +
"<OPTION value='Geneva' style='font-family:Geneva'>Geneva" +
"</SELECT>\
<SELECT name='QBCNTRL3' title='FontSize' onchange='doFormatF(\"FontSize,\"+this.value)'>\
<OPTION value=3>Cỡ chữ\
<OPTION value=7>7\
<OPTION value=6>6\
<OPTION value=5>5\
<OPTION value=4>4\
<OPTION value=3>3\
<OPTION value=2>2\
<OPTION value=1>1\
</OPTION>\
</SELECT>\
</TD></TR>"
str += "<TR valign=middle EVENT>\
<TD bgcolor='#dddddd'>\
<table cellspacing='2' cellpadding='0' border='0' style='border-width:0px;'>\
<tr>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/bgcolor.gif' title='Màu nền' class=vdev onclick='selectBgColor()'>\
</td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/fgcolor.gif' title='Màu chữ' class=vdev onclick='selectFgColor()'>\
</td>\
<td><div class='Divider'></div></td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/cut.gif' title='Cut' class=vdev onclick='doFormatF(\"Cut\")'>\
</td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/copy.gif' title='Copy' class=vdev onclick='doFormatF(\"Copy\")'>\
</td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/paste.gif' title='Paste' class=vdev onclick='doFormatF(\"Paste\")'>\
</td>\
<td><div class='Divider'></div></td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/undo.gif' title='Undo' class=vdev onclick='doFormatF(\"Undo\")'>\
</td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/redo.gif' title='Redo' class=vdev onclick='doFormatF(\"Redo\")'>\
</td>\
<td><div class='Divider'></div></td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/bold.gif' title='Bold' class=vdev onclick='doFormatF(\"Bold\")'>\
</td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/italic.gif' title='Italic' class=vdev onclick='doFormatF(\"Italic\")'>\
</td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/under.gif' title='Underline' class=vdev onclick='doFormatF(\"Underline\")'>\
</td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/strike.gif' title='StrikeThrough' class=vdev onclick='doFormatF(\"StrikeThrough\")'>\
</td>\
<td><div class='Divider'></div></td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/left.gif' title='Left' class=vdev onclick='doFormatF(\"JustifyLeft\")'>\
</td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/center.gif' title='Center' class=vdev onclick='doFormatF(\"JustifyCenter\")'>\
</td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/right.gif' title='Right' class=vdev onclick='doFormatF(\"JustifyRight\")'>\
</td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/justify.gif' title='Justify' class=vdev onclick='doFormatF(\"JustifyFull\")'>\
</td>\
</tr>\
</table>\
</TD></TR>"

if(FULLCTRL)
{
str += "\
<TR class=vdev valign=middle EVENT>\
<TD bgcolor='#dddddd'>\
<table cellspacing='2' cellpadding='0' border='0' style='border-width:0px;'>\
<tr>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/numlist.gif' title='OrderedList' class=vdev onclick='doFormatF(\"InsertOrderedList\")'>\
</td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/bullist.gif' title='UnorderedList' class=vdev onclick='doFormatF(\"InsertUnorderedList\")'>\
</td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/outdent.gif' title='Outdent' class=vdev onclick='doFormatF(\"Outdent\")'>\
</td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/indent.gif' title='Indent' class=vdev onclick='doFormatF(\"Indent\")'>\
</td>\
<td><div class='Divider'></div></td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/superscript.gif' title='SuperScript' class=vdev onclick='doFormatF(\"SuperScript\")'>\
</td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/subscript.gif' title='SubScript' class=vdev onclick='doFormatF(\"SubScript\")'>\
</td>\
<td><div class='Divider'></div></td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/link.gif' title='Create Link' class=vdev onclick='createLink()'>\
</td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/hr.gif' title='HR' class=vdev onclick='doFormatF(\"InsertHorizontalRule\")'>\
</td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/upload.jpg' title='Insert files' class=vdev onclick='InsertFiles()'>\
</td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/image.gif' title='Insert Image' class=vdev onclick='InsertImage()'>\
</td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/Videoclip4.gif' title='Insert Audio-Video' class=vdev onclick='InsertAV()'>\
</td>\
<td><div class='Divider'></div></td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/removeformat.gif' title='Remove HTML' class=vdev onclick='RemoveHTML()'>\
</td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/delformat.gif' title='Delete Format' class=vdev onclick='doFormatF(\"RemoveFormat\")'>\
</td>\
<td><div class='Divider'></div></td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<input type='checkbox' onclick='swapMode()'><font style='font-family: Arial; font-size: 9pt;'>HTML</font>\
</td>\
</tr>\
</table>\
</TD></TR>"

str += "\
<TR class=vdev valign=middle EVENT>\
<TD bgcolor='#dddddd'>\
<table cellspacing='2' cellpadding='0' border='0' style='border-width:0px;'>\
<tr>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/instable.gif' title='InsertTable' class=vdev onclick='insertTable()'>\
</td>\
<td><div class='Divider'></div></td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/inscell.gif' title='InsertCell' class=vdev onclick='insertTableCell()'>\
</td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/delcell.gif' title='DeleteCell' class=vdev onclick='deleteTableCell()'>\
</td>\
<td><div class='Divider'></div></td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/insrow.gif' title='InsertRow' class=vdev onclick='insertTableRow()'>\
</td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/delrow.gif' title='DeleteRow' class=vdev onclick='deleteTableRow()'>\
</td>\
<td><div class='Divider'></div></td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/inscol.gif' title='InsertCol' class=vdev onclick='insertTableCol()'>\
</td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/delcol.gif' title='DeleteCol' class=vdev onclick='deleteTableCol()'>\
</td>\
<td><div class='Divider'></div></td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/mrgcell.gif' title='IncreaseColSpan' class=vdev onclick='morecolSpan()'>\
</td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/spltcell.gif' title='DecreaseColSpan' class=vdev onclick='lesscolSpan()'>\
</td>\
<td><div class='Divider'></div></td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/mrgrow.gif' title='IncreaseRowSpan' class=vdev onclick='morerowSpan()'>\
</td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/spltrow.gif' title='DecreaseRowSpan' class=vdev onclick='lessrowSpan()'>\
</td>\
<td><div class='Divider'></div></td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/div.gif' title='CreateDiv/DivStyle' class=vdev onclick='insertDivLayer()'>\
</td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/divborder.gif' title='DivBorder' class=vdev onclick='editDivBorder()'>\
</td>\
<td onmouseup='mouseUp(this)' onmousedown='mouseDown(this)' onmouseover='mouseOver(this)' onmouseout='mouseOut(this)' class='ButtonNormal'>\
<img src='IURL/divfilter.gif' title='DivFilter' class=vdev onclick='editDivFilter()'>\
</td>\
</tr>\
</table>\
</TD></TR>"
}

 var iurl= QBPATH + '/imgedit'
 var event= "onmousedown='fID=\"" + fid +"\"'"
 str = str.replace(/IURL/g, iurl);
 str = str.replace(/EVENT/g, event);
 return str ;
}



function createEditor(el,id,wi,hi)
{
if(wi=='' || parseInt(wi)<500) wi=625;
  if(hi=='' || parseInt(hi)<100) hi=100;
  
  var hval='';
  if(el.value) hval= el.value;
  hval= hval.replace(/\'/g,"&#39;")
  hval= hval.replace(/&/g,"&amp;")

  var arr = id.split("VDevID")

  var strx = "<iframe id="+id+" style='height:"+hi+"; width:"+wi+"'></iframe>"
  strx += "<input name="+arr[1]+" type=hidden value='"+hval+"'></input>"
  var str="<TABLE border=1 cellspacing=0 cellpadding=1 width="+wi+">";
    str += controlRows(id);
  str +="<tr><td align=center>";
  str += strx+"</td></tr>";
  str += "</TABLE>";

  el.outerHTML= str;

}

function destroyEditor()
{
  var el=document.frames[fID]; 
  if(!el){alert(EDISELECT);return}

  var urlx= QBPATH + '/open/deeditor.html'

  var twidth= 300, theight=140;
  var tposx= (screen.width- twidth)/2
  var tposy= screen.height- theight - 55
  	    	  
  var newWin1=window.open(urlx,"destroy","toolbar=no,width="+ twidth+",height="+ theight+ ",directories=no,status=no,scrollbars=yes,resizable=no, menubar=no")
  newWin1.moveTo(tposx,tposy);
  newWin1.focus()
}



function selectEmoticon()
{ 
  var el=document.frames[fID]
  if(!el){alert(EDISELECT);return}
  el.focus();
  doFormatDialog('emoticon.html','InsertImage',QBPATH)
}

function selectBgColor()
{ 
  doFormatDialog('open/selcolor.html',"BackColor",'')
}


function selectFgColor()
{ 
  doFormatDialog('open/selcolor.html','ForeColor','')
}


function doUploadFile()
{
  var el=document.frames[fID]
  if(!el){alert(EDISELECT);return}
  el.focus()

  var urlx= QBPATH + '/open/upload.html'

  var twidth= 0.8*screen.width, theight=140;
  var tposx= (screen.width- twidth)/2
  var tposy= screen.height- theight - 55
  	    	  
  var newWin1=window.open(urlx,"upload","toolbar=no,width="+ twidth+",height="+ theight+ ",directories=no,status=no,scrollbars=yes,resizable=no, menubar=no")
  newWin1.moveTo(tposx,tposy);
  newWin1.focus()

}


function doEditorOptions()
{
  var el=document.frames[fID]
  if(!el){alert(EDISELECT);return}
  el.focus()

  var urlx= QBPATH + '/open/options.html'

  var twidth= 0.8*screen.width, theight=190;
  var tposx= (screen.width- twidth)/2
  var tposy= screen.height- theight - 55
  	    	  
  var newWin1=window.open(urlx,"options","toolbar=no,width="+ twidth+",height="+ theight+ ",directories=no,status=no,scrollbars=yes,resizable=no, menubar=no")
  newWin1.moveTo(tposx,tposy);
  newWin1.focus()

}


function displayHelp()
{
  var urlx= QBPATH + '/open/edithelp.html'

  var newWin=window.open(urlx,"help","toolbar=no, width=600px,height=400px,directories=no,status=no,scrollbars=yes,resizable=yes,menubar=no;scroll=no")
  newWin.focus()
}


function doExtras()
{
  var el=document.frames[fID]
  if(!el){alert(EDISELECT);return}
  el.focus()

  var urlx= QBPATH + '/open/extras.html'
  var twidth=400, theight=20;
  var tposx= (screen.width- twidth)/2
  var tposy= screen.height- theight - 155
  	    	  
  var newWin1=window.open(urlx,"extras","toolbar=no,width="+ twidth+",height="+ theight+ ",directories=no,status=no,scrollbars=yes,resizable=no, menubar=no")
  newWin1.moveTo(tposx,tposy);
  newWin1.focus()

}



function insertLink(linkurl)
{
  var el=document.frames[fID]
  if(!el && !TXTOBJ){alert(ELESELECT);return}

  if(el)
  {
    el.focus();
    var sel = el.document.selection;
    var strx= "<A href='"+linkurl+"' target=nwin>" + linkurl + "</A>"

    var Range = sel.createRange();
    if(!Range.duplicate) return;
    Range.pasteHTML(strx);
  }
  else 
  {
    TXTOBJ.focus();
    var caret= TXTOBJ.document.selection.createRange()
    TXTOBJ.curword=caret.duplicate();
    var strx= "<A href='"+linkurl+"' target=nwin>" + linkurl + "</A>,"
    doFormat(strx,caret)
  }

}

function editDivBorder()
{
  var el=document.frames[fID]
  if(!el){alert(EDISELECT);return}
  el.focus()
  
  var sel = el.document.selection;
  if(sel==null || sel.type!='Control') {alert(DIVSELECT);return} 

  var Range = sel.createRange();
  if(Range(0).tagName!='DIV') return

  var urlx= QBPATH + '/open/divborder.html'

  var twidth= 0.8*screen.width, theight=215;
  var tposx= (screen.width- twidth)/2
  var tposy= screen.height- theight - 55
  	    	  
  var newWin1=window.open(urlx,"divborder","toolbar=no,width="+ twidth+",height="+ theight+ ",directories=no,status=no,scrollbars=yes,resizable=no, menubar=no")
  newWin1.moveTo(tposx,tposy);
  newWin1.focus()

}




function editDivFilter()
{
  var el=document.frames[fID]
  if(!el){alert(EDISELECT);return}
  el.focus()

  var sel = el.document.selection;
  if(sel==null || sel.type!='Control') {alert(DIVSELECT);return} 

  var Range = sel.createRange();
  if(Range(0).tagName!='DIV') return

  var urlx= QBPATH + '/open/divfilter.html'

  var twidth= 0.8*screen.width, theight=210;
  var tposx= (screen.width- twidth)/2
  var tposy= screen.height- theight - 55
  	    	  
  var newWin1=window.open(urlx,"divfilter","toolbar=no,width="+ twidth+",height="+ theight+ ",directories=no,status=no,scrollbars=yes,resizable=no, menubar=no")
  newWin1.moveTo(tposx,tposy);
  newWin1.focus()

}





function findTextHotKey(forward)
{
  if(!fID && !TXTOBJ){alert(EDISELECT);return}
  if(fID) el=document.frames[fID]
  else el= TXTOBJ
  el.focus();

  var rng = el.document.selection.createRange();
  el.curword=rng.duplicate();

  if(!FWORD && !el.curword.text ){ alert(NOFINDKEY); return }
  else if(el.curword.text)FWORD= el.curword.text

  if(el.curword.text)
   {
     if(forward==1) rng.moveEnd("character", -1 );  
	 else rng.moveStart("character", 1);  
   }

  if(rng.findText(FWORD,100000,FLAGS+forward)==true)
   { rng.select();  rng.scrollIntoView(); return }

  alert(FINDFINISH)
  return

}




function FileDialog()
{
  var urlx= QBPATH + '/open/filedialog.html'

  var twidth= 0.5*screen.width, theight=100;
  var tposx= (screen.width- twidth)/2
  var tposy= screen.height- theight - 55
  	    	  
  var newWin1=window.open(urlx,"fdialog","toolbar=no,width="+ twidth+",height="+ theight+ ",directories=no,status=no,scrollbars=yes,resizable=no, menubar=no")
  newWin1.moveTo(tposx,tposy);
  newWin1.focus()
}




function initDefaultOptions0(fID)
{
   setTimeout("document.frames['"+fID+"'].document.body.style.fontFamily='"+DFFACE+"'",200)
   setTimeout("document.frames['"+fID+"'].document.body.style.fontSize='"+DFSIZE+"'",200)
   setTimeout("document.frames['"+fID+"'].document.body.style.color='"+DCOLOR+"'",200)
   setTimeout("document.frames['"+fID+"'].document.body.style.backgroundColor='"+DBGCOL+"'",200)
   setTimeout("document.frames['"+fID+"'].document.body.style.backgroundImage='url("+DBGIMG+")'",200)
   setTimeout("CSS['"+fID+"']=document.frames['"+fID+"'].document.createStyleSheet('"+DCSS+"')",200)

   FACE[fID]= DFFACE;
   SIZE[fID]= DFSIZE;
   COLOR[fID]= DCOLOR;
   BCOLOR[fID]= DBGCOL;
   BIMAGE[fID]= DBGIMG;
}






function DefaultOptions(linex)
{
  var retArr= new Array('','','','','','','');
  var tempx, strx, objx, idx ;


  // DEFAULT DIV
  var idx= linex.indexOf('ViEtDeVdIvId')
  if(idx>=0) 
	{
	  strx= linex.substring(linex.indexOf('style="')+7,linex.indexOf('">'))

      var atrA= strx.split(";")
	  for(var i=0; i<atrA.length; i++)
		{
		  tempx= atrA[i].split(':')
		  switch(tempx[0].toUpperCase())
		   {
			case "FONT-FAMILY": retArr[0]= tempx[1]; break;
			case "FONT-SIZE": retArr[1]= tempx[1]; break;
			case "BACKGROUND-COLOR": retArr[2]= tempx[1]; break;
			case "COLOR": retArr[3]= tempx[1]; break;
			case "BACKGROUND-IMAGE": if(tempx[2]) tempx[1] += ':'+ tempx[2];
									 retArr[4]= tempx[1].substring(tempx[1].indexOf('url(')+4,tempx[1].indexOf(')') ); 
									 break;
		   }
	    }

	  linex= ""+ />.*<\/div>/i.exec(linex)
      linex= linex.substring(1,linex.length-6)	
    }


   // EXT STYLE
   idx= linex.indexOf('<style>@import url("')
   if( idx>=0 )
    {
	   var strx= linex.substring(idx+20, linex.indexOf('")'))
       retArr[5]= strx
	   linex= linex.substring(0,idx)
    }

   retArr[6]= linex

   return retArr

}





function initDefaultOptions1(linex,fID)
{
  var retArr= new Array();

  retArr= DefaultOptions(linex);

  setTimeout("document.frames['"+fID+"'].document.body.style.fontFamily='"+retArr[0]+"'",200)
  setTimeout("document.frames['"+fID+"'].document.body.style.fontSize='"+retArr[1]+"'",200)
  setTimeout("document.frames['"+fID+"'].document.body.style.backgroundColor='"+retArr[2]+"'",200)
  setTimeout("document.frames['"+fID+"'].document.body.style.color='"+retArr[3]+"'",200)
  setTimeout("document.frames['"+fID+"'].document.body.style.backgroundImage='url("+retArr[4]+")'",200)
  setTimeout("CSS['"+fID+"']=document.frames['"+fID+"'].document.createStyleSheet('"+retArr[5]+"')",200)

  FACE[fID]= retArr[0];
  SIZE[fID]= retArr[1];
  COLOR[fID]= retArr[3];
  BCOLOR[fID]= retArr[2];
  BIMAGE[fID]= retArr[4];

  return retArr[6]

}




function actualize()
{
  var i=0;
  while(document.getElementsByTagName('iframe')[i])
  { 
	setHiddenValue(document.getElementsByTagName('iframe')[i].id) 
	i++
  }
}



function setHiddenValue(fid)
{ 
 if(!fid) return

 var strx= editorContents(fid)

 var idA= fid.split('VDevID')
 if(!idA[0]) return;

 var fobj= document.forms[idA[0]]
 if(!fobj) return;

 var loc=location.href
 loc= loc.substring(0,loc.lastIndexOf('/'))
 if(! /http:\/\//.test(loc) || /http\:\/\/127\.0\.0\.1/.test(loc) || /http\:\/\/localhost/.test(loc))
  {
   loc= loc.replace(/\//g,"\\/")
   loc= loc.replace(/\./g,"\\.")
   var reg= eval("/"+loc+"/g");
   strx= strx.replace(reg,".")
  }

 fobj[idA[1]].value= strx

}	





function doCleanCode(strx,fid) 
{    

  strx = strx.replace(/\r/g,""); 
  strx = strx.replace(/\n>/g,">"); 
  strx = strx.replace(/>\n/g,">"); 

  strx = strx.replace(/\\/g,"&#92;");
  strx = strx.replace(/\'/g,"&#39;")


  // Security
  if(SECURE==1)
	{
	  strx = strx.replace(/<meta/ig, "< meta"); 
	  strx = strx.replace(/&lt;meta/ig, "&lt; meta"); 

	  strx = strx.replace(/<script/ig, "< script"); 
	  strx = strx.replace(/&lt;script/ig, "&lt; script"); 
	  strx = strx.replace(/<\/script/ig, "< /script"); 
	  strx = strx.replace(/&lt;\/script/ig, "&lt; /script"); 

	  strx = strx.replace(/<iframe/ig, "< iframe"); 
	  strx = strx.replace(/&lt;iframe/ig, "&lt; iframe"); 
	  strx = strx.replace(/<\/iframe/ig, "< /iframe"); 
	  strx = strx.replace(/&lt;\/iframe/ig, "&lt; /iframe"); 

	  strx = strx.replace(/<object/ig, "< object"); 
	  strx = strx.replace(/&lt;object/ig, "&lt; object"); 
	  strx = strx.replace(/<\/object/ig, "< /object"); 
	  strx = strx.replace(/&lt;\/object/ig, "&lt; /object"); 

	  strx = strx.replace(/<applet/ig, "< applet"); 
	  strx = strx.replace(/&lt;applet/ig, "&lt; applet"); 
	  strx = strx.replace(/<\/applet/ig, "< /applet"); 
	  strx = strx.replace(/&lt;\/applet/ig, "&lt; /applet"); 

	  strx = strx.replace(/ on/ig, " o&shy;n"); 
	  strx = strx.replace(/script:/ig, "script&shy;:"); 
    }


  var idx= strx.indexOf('ViEtDeVdIvId')
  if( idx>=0 ) strx= strx.substring(strx.indexOf('>')+1,strx.lastIndexOf('</DIV>'))

  idx= strx.indexOf('<style>@import url(')
  if( idx>=0 ) strx= strx.substring(0,idx)
  if(CSS[fid] && CSS[fid].href) strx += '<style>@import url("'+CSS[fid].href+'");</style>';


  var defdiv="" ;
  if(FACE[fid]) defdiv += "; FONT-FAMILY:"+ FACE[fid] 
  if(SIZE[fid]) defdiv += "; FONT-SIZE:"+ SIZE[fid]
  if(COLOR[fid]) defdiv += "; COLOR:"+ COLOR[fid]
  if(BCOLOR[fid])defdiv += "; BACKGROUND-COLOR:"+ BCOLOR[fid]
  if(BIMAGE[fid] && BIMAGE[fid]!='about:blank')
	{
     BIMAGE[fid]= BIMAGE[fid].replace(/\\/g,"/"); 
	 defdiv += "; BACKGROUND-IMAGE:url("+ BIMAGE[fid]+")"
    }
  if(defdiv)
	{
	 defdiv = '<DIV id=ViEtDeVdIvId style="POSITION:Relative' + defdiv + '">'
	 strx = defdiv + strx + "</DIV>"
	}


  // From Valerio Santinelli, PostNuke Developer,(http://www.onemancrew.org)
  // removes all Class attributes on a tag eg. '<p class=asdasd>xxx</p>' returns '<p>xxx</p>'    
     //code = code.replace(/<([\w]+) class=([^ |>]*)([^>]*)/gi, "<$1$3")
  // removes all style attributes eg. '<tag style="asd asdfa aasdfasdf" something else>' returns '<tag something else>'
     //code = code.replace(/<([\w]+) style=\"([^\"]*)\"([^>]*)/gi, "<$1$3")
  // gets rid of all xml stuff... <xml>,<\xml>,<?xml> or <\?xml>
     //code = code.replace(/<]>/gi">\\?\??xml[^>]>/gi, "")
  // get rid of ugly colon tags <a:b> or </a:b>
     //code = code.replace(/<\/?\w+:[^>]*>/gi, "")
  // removes all empty <p> tags
     strx = strx.replace(/<p([^>])*>(&nbsp;)*\s*<\/p>/gi,"")
  // removes all empty span tags
     strx = strx.replace(/<span([^>])*>(&nbsp;)*\s*<\/span>/gi,"")
  return strx
}





//////////////////////////////////////////////////////////////////////
function addEventToObj()
{
  // addEventListener -> all Textarea
  var oArr= document.getElementsByTagName("textarea")
  var i=-1;
  while(oArr[++i])
   {
	 oArr[i].attachEvent("onmousedown",doMDown)
	 oArr[i].attachEvent("onmouseup",doMUp)
	 oArr[i].attachEvent("onkeydown",doKDown)
   }

  // addEventListener -> all Input
  oArr= document.getElementsByTagName("input")
  i=-1
  while(oArr[++i])
   {
	 oArr[i].attachEvent("onmousedown",doMDown)
	 oArr[i].attachEvent("onmouseup",doMUp)
	 if(oArr[i].type!="text") continue
	 oArr[i].attachEvent("onkeydown",doKDown)
   }
}


addEventToObj();




function editorContents(fid)
{
  var el= document.frames[fid]
  if(!el)return

  var strx, strx1;
  if(format[fid]=="HTML")
	{
	  if(curTD)
	   { 
   	     curTD.runtimeStyle.backgroundColor = "";
		 curTD.runtimeStyle.color = "";
		 curTD=null 
		 curTB.runtimeStyle.backgroundColor = "";
		 curTB.runtimeStyle.color = "";
		 curTB=null 
	   }
	  strx= el.document.body.innerHTML
	  strx1= el.document.body.innerText
	}
  else
	{
	  strx = el.document.body.innerText
	  strx1=el.document.body.innerHTML
    }
  if(strx1=='' && strx.indexOf('<IMG')<0 && strx.indexOf('<HR')<0 ) return ''


  strx = doCleanCode(strx,fid);

  return strx
}




function setTextStyle(tstyle)
{
  if(!tstyle) return

  var el= document.frames[fID]
  if(!el){alert(EDISELECT);return}
  el.focus()

  var edit=el.document; 

  var sArr= tstyle.split(",") // FontName,ForeColor,HiliteColor,FontSize,Italic

  edit.execCommand("RemoveFormat",false,false) 

  if(sArr[0]) edit.execCommand("FontName",false,sArr[0]) 
  if(sArr[1]) edit.execCommand("ForeColor",false,sArr[1]) 
  if(sArr[2]) edit.execCommand("BackColor",false,sArr[2])
  if(sArr[3]) edit.execCommand("FontSize",false,sArr[3])
  if(sArr[4]) edit.execCommand("Italic",false,sArr[4])

}



function doFormatF(arr)
{
  var el= document.frames[fID]
  if(!el){alert(EDISELECT);return}
  el.focus()

  var cmd = new Array();
  cmd = arr.split(',')

  var edit=el.document; 
  if(cmd[0]=='formatBlock')
   {
    edit.execCommand(cmd[0],false,"<"+cmd[1]+">");
    if(cmd[1]=='PRE' && format[fID]=="HTML") swapMode();
   }
  else if(cmd[0]=='InsertImage' && !cmd[1] )
   {
    alert(IMAGESRC); 
    edit.execCommand(cmd[0],true,null) 
   }
  else if(cmd[1]!=null) edit.execCommand(cmd[0],false,cmd[1]) 
  else edit.execCommand(cmd[0],false,null)

}



function insertImageSimple(el,cmd)
{
  var html= '<img src="' + cmd +'">'
  insertHTML(el,html)
}





function swapCharCode()
{
 var el= document.frames[fID]
 if(!el){alert(EDISELECT);return}
 el.focus()

 var eStyle= el.document.body.style;
 var strx;
 if(format[fID]=="HTML")
 {
  swapMode();
  strx= el.document.body.innerText
  format[fID]="Text"
 }
 else if(viewm[fID]==0)
 {
  strx= el.document.body.innerHTML
  strx= strx.replace(/\&amp;#/g,"&#")
  el.document.body.innerHTML= strx
  viewm[fID]=1 - viewm[fID]
  return
 }
 else
 {
  strx= el.document.body.innerText
 }

 if(viewm[fID]) strx=toUnicode(strx)
 
 el.document.body.innerText=strx

 viewm[fID]=1 - viewm[fID]
}


function toUnicode(str1)
{
  var code, str2 , j=0;
  var len
  while(j<2)
   {
	len=str1.length
	str2=''
	for(var i=0;i<len;i++) 
	 {
      code=str1.charCodeAt(i);
      if(code<128) continue;
      str2 +=str1.substring(0,i) + '&#' + code + ';'
      str1=str1.substring(i+1,str1.length)
      len=str1.length
      i=0
     }
    str1=str2+str1
    j++;
   }
  return str1;
}



function swapMode()
{
 var el= document.frames[fID]
 if(!el){alert(EDISELECT);return}
 el.focus()


 var MARK= "ViEtDeVtRiCk"
 var selType=el.document.selection.type

 if(selType!="Control")
 {
   var caret=el.document.selection.createRange();
   el.curword=caret.duplicate();
   var selwrd= el.curword.text
   el.curword.text = selwrd + MARK;
 }

 var eStyle= el.document.body.style

	 
 if(format[fID]=="HTML")
 {
  FACE[fID]= eStyle.fontFamily
  SIZE[fID]= eStyle.fontSize
  COLOR[fID]= eStyle.color
  BCOLOR[fID]= eStyle.backgroundColor
  BIMAGE[fID]= eStyle.backgroundImage
  BIMAGE[fID]= BIMAGE[fID].substring( BIMAGE[fID].indexOf('(')+1,BIMAGE[fID].indexOf(')') )

  eStyle.fontFamily="";
  eStyle.fontSize="12pt"
  eStyle.fontStyle="normal"
  eStyle.color="black"
  eStyle.backgroundColor="#e0e0f0"
  eStyle.backgroundImage=''
  var innerHTML= el.document.body.innerHTML
  var reg= eval("/"+MARK+"/ig");
  var res= innerHTML.match(reg);
  if(res)
   for(var i=0; i<res.length-1; i++)
	 innerHTML= innerHTML.replace(res[i],"") 

  el.document.body.innerText= innerHTML;
  format[fID]="Text"
 }
 else
 {
  eStyle.fontFamily= FACE[fID]
  eStyle.fontSize= SIZE[fID]
  eStyle.color= COLOR[fID]
  eStyle.backgroundColor= BCOLOR[fID]
  eStyle.backgroundImage= "url(" + BIMAGE[fID] + ")"

  var temp=el.document.body.innerText
  el.document.body.innerHTML= temp;

  format[fID]="HTML"
  viewm[fID]=1

  // addeventlistener for table-cell
  var tdA= el.document.getElementsByTagName('td')
  for(var i=0; i<tdA.length;i++)
   { tdA[i].attachEvent("onclick", clickTD) }

 }


 if(selType!="Control")
 {
  caret = el.document.selection.createRange();
  var found= caret.findText(MARK,100000,5) // backward
  if(found==false) 
   found= caret.findText(MARK,100000,4) // foreward

  if(found==false && format[fID]=="HTML") 
   {
     var strx= el.document.body.innerHTML
	 strx= strx.replace(/ViEtDeVtRiCk/ig,"");
	 el.document.body.innerHTML= strx
	 return;
   }

  caret.select();
  el.curword=caret.duplicate();
  el.curword.text = '' ;  // erase trick selection 

  if(selwrd!="") caret.findText(selwrd,100000,5); // real selection
  caret.select();  caret.scrollIntoView(); 
 }

}




function selectAll()
{ 
  var el= document.frames[fID]
  if(!el){alert(EDISELECT);return}
  el.focus()
  
  var s=el.document.body.createTextRange()
  s.execCommand('SelectAll',false,null)
}





function highLight(key)
{
  function doDefFormat()
   {
     var el= document.frames[fID]
     el.focus();
     var rng = el.document.selection.createRange();
     rng.moveEnd("character", 1);
     rng.select();
     el.curword=rng.duplicate();
     if(el.curword.text=='') doFormatF('RemoveFormat'); 
     else
     {
      rng.moveEnd("character", -1);
      rng.select();
      doFormatF('ForeColor,'); doFormatF('BackColor,'); 
     }
    }

  switch(key)
	{  
	  case 48: doDefFormat(); break; // ctrl+0  no highlight
	  case 49: doFormatF('ForeColor,red'); break; // ctrl+1
	  case 50: doFormatF('ForeColor,green'); break; // ctrl+2
	  case 51: doFormatF('ForeColor,blue'); break; // ctrl+3
      case 52: doFormatF('ForeColor,#00AAFF'); break; // ctrl+4
      case 53: doFormatF('ForeColor,magenta'); break; // ctrl+5
	  case 54: doFormatF('BackColor,yellow'); doFormatF('ForeColor,black'); break; // ctrl+6
	  case 55: doFormatF('BackColor,cyan'); doFormatF('ForeColor,black'); break; // ctrl+7
	  case 56: doFormatF('BackColor,#00FF00'); doFormatF('ForeColor,black'); break; // ctrl+8
	  case 57: doFormatF('BackColor,#FF00AA'); doFormatF('ForeColor,white'); break; // ctrl+9
    }
}




function FKDown()
{
  var el= document.frames[fID]
  var event= el.event

  if(!el ||!event){alert(EDISELECT);return}
  if(event.altKey) return;

  var key= event.keyCode
  var shft= event.shiftKey
  var ctrl= event.ctrlKey


  if(RETURNNL && !shft && key==13){ insertNewline(el); return false }
  else if(RETURNNL && key==13){ insertNewParagraph(el); return false }

  if(ctrl && key==71){ findText(); return false }  // ctrl+G search
  else if(ctrl && key==75){ findTextHotKey(0); return false } // ctrl+K  search forward
  else if(ctrl && key==74){ findTextHotKey(1); return false } // ctrl+J  search backward 
  else if(ctrl && key==83 && SYMBOLE!=''){ SmartcardData(); return false } // ctrl+S content rewrite
  else if(ctrl && key==84){ swapMode(); return false } // ctrl+T swapMode
  else if(ctrl && (key>=48 && key<=57)){ highLight(key); return false } // ctrl+1 Highlight

}




function insertHTML(el,html)
{
  var sel = el.document.selection;
  if(sel.type=="Control") return 

  var Range = sel.createRange();
  if(!Range.duplicate) return;
  var wrd='' ;
  el.curword=Range.duplicate();
  wrd= el.curword.text;

  var Range = sel.createRange();
  if(!Range.duplicate) return;
  Range.pasteHTML(html);
}



function insertDivLayer()
{
  var el= document.frames[fID]
  if(!el){alert(EDISELECT);return}
  el.focus()
  
  var sel = el.document.selection;
  if(sel==null) return

  var Range = sel.createRange();
  var wrd='' ;

  if(sel.type!="Control")
  {
  	if(!Range.duplicate) return;
  	el.curword=Range.duplicate();
  	wrd= el.curword.text;
	if(wrd=='') wrd="I'm a DIV-Layer. Select me and click the button once more to change properties. Or doubleclick me to change the text."
	var arr= "<DIV style='position:relative; width:150px; height:100px; font-family:Arial; font-size:12px; background-color:#f0fdd0; border:1 solid'>"+ wrd + "</DIV>" ;
	Range.pasteHTML(arr);
	return
  }  

  if(Range(0).tagName!='DIV') return

  var urlx = QBPATH + '/open/divstyle.html'

  var twidth= 0.8*screen.width, theight=190;
  var tposx= (screen.width- twidth)/2
  var tposy= screen.height- theight - 55
  	    	  
  var newWin1=window.open(urlx,"divstyle","toolbar=no,width="+ twidth+",height="+ theight+ ",directories=no,status=no,scrollbars=yes,resizable=no, menubar=no")
  newWin1.moveTo(tposx,tposy);
  newWin1.focus()

}




function formatDialog()
{
  TXTOBJ.focus();
  var caret=TXTOBJ.document.selection.createRange()
  TXTOBJ.curword=caret.duplicate();
  
  var y = screen.height -parseInt('27em')*14 - 30 
  var feature = "font-family:Arial;font-size:10pt;dialogWidth:30em;dialogHeight:27em;dialogTop:"+y
      feature+= ";edge:sunken;help:no;status:no"

  var dialog= QBPATH+'/open/dialog.html'
  var arr= showModalDialog(dialog, "", feature); //////////////////////////////////////////
  if(arr==null) return ;

  if(arr=='VISUAL'){ changetoIframeEditor(TXTOBJ); userAddInit() }
  else doFormat(arr,caret)
}




function createLink()
{
  var el= document.frames[fID]
  if(!el){alert(EDISELECT);return}
  el.focus()

  var urlx= QBPATH + '/open/createlink.html'

  var arr=showModalDialog(urlx, el, 
	  "font-family:Verdana;font-size:12;dialogWidth:30em;dialogHeight:14em; edge:sunken;help:no;status:no");

}




function doFormatDialog(file,cmd,arg)
{ 
  var urlx= QBPATH + '/' + file

  var el=document.frames[fID];
  if(!el){alert(EDISELECT);return}

  var arr=showModalDialog(urlx, arg, "font-family:Verdana;font-size:12;dialogWidth:30em;dialogHeight:30em; edge:sunken;help:no;status:no");
  if(arr !=null) doFormatF(cmd+','+arr)
}



function characters()
{
  var el=document.frames[fID];
  if(!el){alert(EDISELECT);return}
  el.focus();

  var sel = el.document.selection;
  if(sel.type=="Control") return 

  var urlx= QBPATH + '/open/selchar.html'

  var arr=showModalDialog(urlx, '', "font-family:Verdana;font-size:12;dialogWidth:30em;dialogHeight:34em; edge:sunken;help:no;status:no");
  if(arr==null) return

  var arrA = arr.split(';QuIcKbUiLd;')

  var strx= "<FONT FACE='" + arrA[0] + "'>" + arrA[1] + "</FONT>"

  var Range = sel.createRange();
  if(!Range.duplicate) return;
  Range.pasteHTML(strx);

}




///////////////////////////////////////////////////////////////////////
if(USETABLE) document.writeln('<script src="'+QBPATH+'/js/tabedit.js"></script>');
if(RETURNNL) document.writeln('<script src="'+QBPATH+'/js/returnnl.js"></script>');
//if(UNICODE) document.writeln('<script src="'+QBPATH+'/js/unicode.js"></script>');
document.writeln('<script src="'+QBPATH+'/js/recover.js"></script>');



// VISUAL=0 : Textarea to Editor after confirmation
// VISUAL=1 : all Textarea to Editor
// VISUAL=2 : change only specific textarea
// VISUAL=3 : all Iframe to Editor
// VISUAL=4 : some specific iframes 
// VISUAL=other : no Visual-Editor, only use Rightmouse-Control
switch(VISUAL)
{
  case 1: changeAllTextareaToEditors(); break;
  case 2: changetoIframeEditor(document.forms[xxx].yyy); break;// please replace xxx=formIndex and yyy=textareaName
  case 3: changeAllIframeToEditors(); break;
  case 4: changeIframeToEditor('contents2'); break;//please replace contents.. = frame id
}

function doMDown()
{
// var el;
// el= event.srcElement
 
// var button= event.button

// if(el.type=='text' || el.type=='textarea')
//{
//    TXTOBJ=el; fID=''
//    if(button>1 && POPWIN==1){ formatDialog();}
//   }
}

function doMUp()
{
 el= event.srcElement
 if(!el.type) return
 var fidx= fID

 if(el.type!='text'&&el.type!='textarea'&&el.type!='password'&&el.type!='file')
  {
   if(!el.name || el.name.substring(0,7)!='QBCNTRL')
    { 
     actualize();
     if(el.type != 'select-one' && el.type != 'select-multiple') el.focus(); 
    }
    fID= fidx
    return
  }

 var visual=''
 if(typeof(ASKED)=="undefined" && el.type=='textarea' && VISUAL==0)
  { visual=confirm(VISMODE); if(!visual) ASKED=1; }
 	 
 if(visual){ changetoIframeEditor(el); userAddInit() }
 else{ TXTOBJ= el; fID=null }
}

/* for use if user you to add function after creating editor */
function userAddInit(){};



////////////////////////////
function doKDown()
{
  var ctrl= event.ctrlKey
  if(!ctrl) return;

  var el=event.srcElement 
  if(el.type!='text' && el.type!='textarea') return
  TXTOBJ=el; fID='';

  var key= event.keyCode
  if(ctrl && key==71) { findText(); return false }  // ctrl+G search
  else if(ctrl && key==75){ findTextHotKey(0); return false } // ctrl+K  search forward
  else if(ctrl && key==74){ findTextHotKey(1); return false } // ctrl+J  search backward 
  else if(ctrl && key==83 && SYMBOLE!=''){ SmartcardData(); return false } // ctrl+S content rewrite
 
}



function findText()
{
  if(!fID && !TXTOBJ){alert(EDISELECT);return}
  if(fID) document.frames[fID].focus()
  else TXTOBJ.focus()

  var urlx= QBPATH + '/open/dfindtext.html'

  var newWin=window.open(urlx,"find","toolbar=no, width=350px,height=220px,directories=no,status=no,scrollbars=yes,resizable=yes,menubar=no;scroll=no")
  newWin.moveTo(screen.width-500,50);
  newWin.focus()
}



