<div id="p1" class="intro"></div>
<div id="p2" class="intro"></div>
<div id="p3" class="intro"></div>
<div id="p4" class="intro"></div>
<div id="p5" class="intro"></div>
<div id="p6" class="intro"></div>
<div id="p7" class="intro"></div>
<div id="p8" class="intro"></div>
<div id="p9" class="intro"></div>
<div id="p10" class="intro"></div>
<div id="p11" class="intro"></div>
<div id="p12" class="intro"></div>
<div id="p13" class="intro"></div>
<div id="p14" class="intro"></div>
<div id="p15" class="intro"></div>
<div id="p16" class="intro"></div>
<div id="p17" class="intro"></div>
<div id="p18" class="intro"></div>
<div id="p19" class="intro"></div>
<div id="p20" class="intro"></div>
<div id="p21" class="intro"></div>
<div id="p22" class="intro"></div>
<div id="p23" class="intro"></div>
<div id="p24" class="intro"></div>
<div id="p25" class="intro"></div>

<script>
var espeed=20
var counter=1
var temp=new Array()
var temp2=new Array()
var ns4=document.layers?1:0
var ie4=document.all?1:0
var ns6=document.getElementById&&!document.all?1:0
if (ns4)
{
	for (i=1;i<=25;i++)
	{
		temp[i]=eval("document.p"+i+".clip")
		temp2[i]=eval("document.p"+i)
		temp[i].width=window.innerWidth/5
		temp[i].height=window.innerHeight/5
}
for (i=1;i<=5;i++)
	temp2[i].left=(i-1)*temp[i].width
for (i=6;i<=10;i++)
{
	temp2[i].left=(i-6)*temp[i].width
	temp2[i].top=temp[i].height
}
for (i=11;i<=15;i++)
{
	temp2[i].left=(i-11)*temp[i].width
	temp2[i].top=2*temp[i].height
}

for (i=16;i<=20;i++)
{
	temp2[i].left=(i-16)*temp[i].width
	temp2[i].top=3*temp[i].height
}

for (i=21;i<=25;i++)
{
	temp2[i].left=(i-21)*temp[i].width
	temp2[i].top=4*temp[i].height
}

}
function erasecontainerns()
{
	window.scrollTo(0,0)
	var whichcontainer=Math.round(Math.random()*25)
	if (whichcontainer==0)
	whichcontainer=1
	if (temp2[whichcontainer].visibility!="hide")
		temp2[whichcontainer].visibility="hide"
	else
	{
		while (temp2[whichcontainer].visibility=="hide")
		{
			whichcontainer=Math.round(Math.random()*25)
			if (whichcontainer==0)
				whichcontainer=1
		}
		temp2[whichcontainer].visibility="hide"
	}
	if (counter==25)
	clearInterval(beginerase)
	counter++
	espeed-=10
}

if (ie4||ns6)
{
	var containerwidth=ns6?parseInt(window.innerWidth)/5-3 : parseInt(document.body.clientWidth/5)
	var containerheight=ns6?parseInt(window.innerHeight)/5-2 : parseInt(document.body.offsetHeight/5)
	for (i=1;i<=25;i++)
	{
		temp[i]=ns6?document.getElementById("p"+i).style : eval("document.all.p"+i+".style")
		temp[i].width=containerwidth	
		temp[i].height=containerheight
	}
	for (i=1;i<=5;i++)
	temp[i].left=(i-1)*containerwidth

	for (i=6;i<=10;i++)
	{
		temp[i].left=(i-6)*containerwidth
		temp[i].top=containerheight
	}

	for (i=11;i<=15;i++)
	{
		temp[i].left=(i-11)*containerwidth
		temp[i].top=2*containerheight
	}

	for (i=16;i<=20;i++)
	{
		temp[i].left=(i-16)*containerwidth
		temp[i].top=3*containerheight
	}

	for (i=21;i<=25;i++)
	{
		temp[i].left=(i-21)*containerwidth
		temp[i].top=4*containerheight
	}
}

function erasecontainerie()
{
	window.scrollTo(0,0)
	var whichcontainer=Math.round(Math.random()*25)
	if (whichcontainer==0)
		whichcontainer=1
	if (temp[whichcontainer].visibility!="hidden")
		temp[whichcontainer].visibility="hidden"
	else
	{
		while (temp[whichcontainer].visibility=="hidden")
		{
			whichcontainer=Math.round(Math.random()*25)
			if (whichcontainer==0)
				whichcontainer=1
		}
	temp[whichcontainer].visibility="hidden"
	}

	if (counter==25)
	{
		clearInterval(beginerase)
		if (ns6)
		{
			for (i=1;i<26;i++)
				temp[i].display="none"
		}
	}	
	counter++
	espeed-=10
}

if (ns4)
	beginerase=setInterval("erasecontainerns()",espeed)
else if (ie4||ns6)
{
	beginerase=setInterval("erasecontainerie()",espeed)
}

</script>