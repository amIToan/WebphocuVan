function insertNewline(obj)
{
  if(obj.document.selection.type=='Control') return;
  var Range = obj.document.selection.createRange();
  if(!Range.duplicate) return;
  Range.pasteHTML('<br>.');
  Range.findText('.',10,5)
  Range.select();
  obj.curword=Range.duplicate();
  obj.curword.text = ''; 
  Range.select();
}

function insertNewParagraph(obj)
{
  if(obj.document.selection.type=='Control') return;
  var Range = obj.document.selection.createRange();
  if(!Range.duplicate) return;

  var parent=Range.parentElement()
  var tagLI= 0
  while(parent && parent.tagName!='BODY')
  {
	if(parent.tagName=='LI'){ tagLI=1; break }
	parent= parent.parentElement
  }

  if(tagLI) Range.pasteHTML('<LI>.');
  else Range.pasteHTML('<P>.');

  Range.findText('.',10,5)
  Range.select();
  obj.curword=Range.duplicate();
  obj.curword.text = '' ;
  Range.select();

}
function RemoveHTML()
{
  var el=document.frames[fID]

  if(!el)return

  var strx, strtmp;
  if(format[fID]=="HTML")
	{
	  strx= el.document.body.innerHTML
	}
  else
	{
	  strx = el.document.body.innerText
    }
  var thedong,themo;
  
  var themo = strx.indexOf("<");

  while (themo != -1) {
  	strtmp=strx.substr(themo, 8);
  	strtmp=strtmp.toLowerCase();

  	if ( (strtmp.match("<img")=="<img") || (strtmp.match("<br")=="<br") || (strtmp.match("<table")=="<table") || (strtmp.match("</table")=="</table") || (strtmp.match("<tr")=="<tr") || (strtmp.match("</tr")=="</tr") || (strtmp.match("<td")=="<td") || (strtmp.match("</td")=="</td") || (strtmp.match("<a ")=="<a ") || (strtmp.match("</a")=="</a") )
  	{
    	themo = strx.indexOf("<", themo + 1);
    }
    else
    {
    	thedong = strx.indexOf(">", themo + 1)+1;
    	if (strtmp.match("</p>")=="</p>")
    	{
    		strx=strx.replace(strx.substring(themo, thedong),"<br>");
    	}
    	else
    	{
    		strx=strx.replace(strx.substring(themo, thedong),"");
		}
    	themo = strx.indexOf("<",themo);
    }
  }
  
  el.document.body.innerHTML=strx;
  return;
}