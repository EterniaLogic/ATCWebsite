do_cookie();
var II=0;
function cookie_cleanup()
{
	cookiedata1=getCookie("syst_windows00192").replace(' ','%20').split("%20");
		var dat1 = "";
		for(var o1=0;o1<=cookiedata1.length-1;o1++)
		{
			d=" ";
			if(o1 == cookiedata1.length-1) d=""; 
			if(cookiedata1[o1] != "" && dat1.search(cookiedata1[o1]) == -1 && cookiedata1[o1] != "undefined")
				dat1+=cookiedata1[o1]+d;
		}
		//alert("|"+dat1.substring(dat1.length-1,dat1.length)+"|");
		if(dat1.substring(dat1.length-1,dat1.length) == " ") dat1=dat1.substring(0,dat1.length-1);
		setCookie("syst_windows00192",dat1.replace(/  /g,' ').replace(/%20/g," "),365*50);
}
function do_cookie()
{
	II++;
	//clearTimeout(t);
	//if(checkCookie("syst_windows00192"))
	//{
		//throw getCookie("syst_windows00192");
		
		//cookie cleanup
		cookie_cleanup();
		
		//setCookie("syst_windows00192","",365*50);
		try{
			cookie=getCookie("syst_windows00192")
			
			cookiedata=cookie.replace(/  /g,' ').replace(/ /g,'|').replace(/%20/g,'|').split("|");
			//alert(cookiedata); 
			//alert(getCookie("syst_windows00192"));
			for(var o1=0;o1<=cookiedata.length-1;o1++)
			{
				//document.write(o);
				try{
					if(cookiedata[o1] != "") {
						//pausecomp(1000);
						//alert("|"+cookiedata[o1]+"|");
						var dat = cookiedata[o1].replace(/ /g,'');
						//alert("|"+dat+"|");
						//document.write("|"+dat+"|");
						var element1 = document.getElementById(dat);
						//element1.parentNode.head.data.mini.mini.id="minimize_down";
						element1.style.visibility = "hidden";
						element1.style.position="absolute";
						//element1.style.left="-200";
						element1.parentNode.style.height="22"; 
						//element1.parentNode.style.visibility="hidden";
						//element1.parentNode.childNodes[3].height="22";
						//do the following w/o needing to find nodes...
						y=element1.parentNode.childNodes[1].childNodes[0].childNodes[0];
						y.childNodes[y.childNodes.length-1].childNodes[0].id = "minimize_down";
						//document.write(element1.parentNode.childNodes[1].childNodes[0].childNodes[0].childNodes[1].childNodes[0].id);
					}
				}
				catch(err){/*document.write(err);*/}
			}
		}
		catch(err){/*document.write(err);*/}
	//}
	
}		


//minimize script
//method to change an elements parent
function setParent(el, newParent) {    newParent.appendChild(el);}
function do_mini(this1,id)
{
	//removes the visible attribute...
	element = document.getElementById(id);
	
	//determine if it is visible or not...
	vis=false;
	if(element.style.visibility == "visible") {element.style.posititon="absolute";element.style.left="-200";element.style.height="0px";element.style.visibility = "hidden";this1.id='minimize_down_over';element.style.height="0";}
	else {vis=true;element.style.posititon="static";element.style.visibility = "visible";this1.id='minimize';element.style.height="auto";}
	//throw id;
	try{
		cookiedata=getCookie("syst_windows00192").replace(/ /g,'%20').split("%20");
		var datar = "";
		//remove id if it is there...
		for(o=0;o<=cookiedata.length;o++)
		{
			d=" ";
			//if(o == cookiedata.length-1) d="";                     //no undefined				no duplicates
			if(cookiedata[o] != id && cookiedata[o] != "" && cookiedata[o] != null && datar.search(cookiedata[o]) == -1) 
				datar+=cookiedata[o]+d;
		}
		if(!vis && datar.search(cookiedata[o]) == -1){
			setCookie("syst_windows00192",datar+id+" ",365*50); }
		else{
			//set to previous data...
			if(datar == "undefined ")
			{
				setCookie("syst_windows00192",id+" ",365*50);
			}
			else{
			setCookie("syst_windows00192",datar,365*50);
			//document.write("set0");
			}
		}
	}
	catch(err)
	{
		//throw err;
		if(id != "undefined"){
			setCookie("syst_windows00192",id+" ",365*50);
			document.write("set2");
			}
	}
	cookie_cleanup();
	return false;
}
function mini_mouseout(this1)
{
	if(this1.id != 'minimize_down' && this1.id != 'minimize_down_over') this1.id='minimize'; 
	else this1.id='minimize_down'; 
}
function mini_mousein(this1)
{
	if(this1.id != 'minimize_down' && this1.id != 'minimize_down_over') this1.id='minimize_over'; 
	else this1.id='minimize_down_over'; 
}

//drag & drop box [http://www.dunnbypaul.net/js_mouse/]
var mousex = 0;
var mousey = 0;
var grabx = 0;
var graby = 0;
var orix = 0;
var oriy = 0;
var elex = 0;
var eley = 0;
var algor = 0;

var dragobj = null;

function falsefunc() { return false; } // used to block cascading events

function init()
{
  document.onmousemove = update; // update(event) implied on NS, update(null) implied on IE
  update();
  do_cookie();
}

function getMouseXY(e) // works on IE6,FF,Moz,Opera7
{ 
  if (!e) e = window.event; // works on IE, but not NS (we rely on NS passing us the event)

  if (e)
  { 
	if (e.pageX || e.pageY)
	{ // this doesn't work on IE6!! (works on FF,Moz,Opera7)
	  mousex = e.pageX;
	  mousey = e.pageY;
	  algor = '[e.pageX]';
	  if (e.clientX || e.clientY) algor += ' [e.clientX] '
	}
	else if (e.clientX || e.clientY)
	{ // works on IE6,FF,Moz,Opera7
	  mousex = e.clientX + document.body.scrollLeft;
	  mousey = e.clientY + document.body.scrollTop;
	  algor = '[e.clientX]';
	  if (e.pageX || e.pageY) algor += ' [e.pageX] '
	}  
  }
}

function update(e)
{
  getMouseXY(e); // NS is passing (event), while IE is passing (null)
  try{
  document.getElementById('span_browser').innerHTML = navigator.appName;
  document.getElementById('span_winevent').innerHTML = window.event ? window.event.type : '(na)';
  document.getElementById('span_autevent').innerHTML = e ? e.type : '(na)';
  document.getElementById('span_mousex').innerHTML = mousex;
  document.getElementById('span_mousey').innerHTML = mousey;
  document.getElementById('span_grabx').innerHTML = grabx;
  document.getElementById('span_graby').innerHTML = graby;
  document.getElementById('span_orix').innerHTML = orix;
  document.getElementById('span_oriy').innerHTML = oriy;
  document.getElementById('span_elex').innerHTML = elex;
  document.getElementById('span_eley').innerHTML = eley;
  document.getElementById('span_algor').innerHTML = algor;
  document.getElementById('span_dragobj').innerHTML = dragobj ? (dragobj.id ? dragobj.id : 'unnamed object') : '(null)';
  }
  catch(err)
  {
  
  }
}

function grab(context)
{
  document.onmousedown = falsefunc; // in NS this prevents cascading of events, thus disabling text selection
  dragobj = context;
  dragobj.style.zIndex = 10; // move it to the top
  document.onmousemove = drag;
  document.onmouseup = drop;
  grabx = mousex;
  graby = mousey;
  elex = orix = dragobj.offsetLeft;
  eley = oriy = dragobj.offsetTop;
  update();
}

function drag(e) // parameter passing is important for NS family 
{
  if (dragobj)
  {
	elex = orix + (mousex-grabx);
	eley = oriy + (mousey-graby);
	dragobj.style.position = "absolute";
	dragobj.style.left = (elex).toString(10) + 'px';
	dragobj.style.top  = (eley).toString(10) + 'px';
  }
  update(e);
  return false; // in IE this prevents cascading of events, thus text selection is disabled
}

function drop()
{
  if (dragobj)
  {
	dragobj.style.zIndex = 0;
	dragobj = null;
  }
  update();
  document.onmousemove = update;
  document.onmouseup = null;
  document.onmousedown = null;   // re-enables text selection on NS
}
function getCookie(c_name)
{
if (document.cookie.length>0)
  {
  c_start=document.cookie.indexOf(c_name + "=");
  if (c_start!=-1)
	{
	c_start=c_start + c_name.length+1;
	c_end=document.cookie.indexOf(";",c_start);
	if (c_end==-1) c_end=document.cookie.length;
	return unescape(document.cookie.substring(c_start,c_end));
	}
  }
return "";
}

function setCookie(c_name,value,expiredays)
{
var exdate=new Date();
exdate.setDate(exdate.getDate()+expiredays);
document.cookie=c_name+ "=" +escape(value)+
((expiredays==null) ? "" : ";expires="+exdate.toGMTString());
}
function HLROW(this1,color,color1)
{
	this1.style.backgroundColor=color;
}
//do_cookie();